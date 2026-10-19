"""Store submission control for the release pipeline.

The green run uploads the verified ``.msixbundle`` as a draft submission and
stops: nothing a customer can see changes, and certification has not started.
The publication pass, which runs only when the release really goes out, sets
``targetPublishMode: "Immediate"`` on that draft and commits it, so the Store
build goes live as soon as it passes certification. The submission API has no
call that releases a submission once it is committed and held, which is why
the commit waits for publication instead of happening at green with Manual.
Both entry points are safe to rerun.

Green run — ``python -m scripts.releases.store submit <package>`` (Windows
runner, ``msstore`` CLI configured beforehand by
``microsoft/microsoft-store-apppublisher`` + ``msstore reconfigure``):

1. ``msstore submission get <productId>`` — the current submission as JSON.
2. ``msstore submission delete <productId> --no-confirm`` — only for a draft
   (``PendingCommit``) or a failed submission. A committed submission that is
   still in certification or publishing belongs to an earlier release going
   live; ``msstore publish`` would delete it, so ``submit`` refuses instead.
3. ``msstore publish <package> --appId <productId> --noCommit`` — uploads the
   package and leaves the submission as a draft (``--noCommit`` / ``-nc``,
   "Disables committing the submission, keeping it in draft state").

Publication pass — ``scripts.releases.store.release(...)`` (Ubuntu runner,
Partner Center submission REST API, stdlib ``urllib`` only):

1. ``POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`` with
   ``grant_type=client_credentials`` and scope
   ``https://manage.devcenter.microsoft.com/.default`` (the same MS_STORE_*
   credentials the CLI uses).
2. ``GET https://manage.devcenter.microsoft.com/v1.0/my/applications/{id}`` —
   ``pendingApplicationSubmission.id`` names the draft.
3. ``GET .../submissions/{id}`` — the draft must carry this release's Store
   bundle (``applicationPackages[].fileName``); a draft from a newer green run
   is left alone.
4. ``PUT .../submissions/{id}`` with ``targetPublishMode: "Immediate"``, then
   ``POST .../submissions/{id}/commit``; certification starts and the build
   goes live when it passes. A submission that is already committed or live is
   a no-op.

Sources for the commands and fields above:
- msstore CLI commands and options (``submission get/delete --no-confirm``,
  ``publish --noCommit``):
  https://learn.microsoft.com/en-us/windows/apps/publish/msstore-dev-cli/commands
- CLI CI/CD setup (``msstore reconfigure --tenantId --sellerId --clientId
  --clientSecret`` on the runner):
  https://learn.microsoft.com/en-us/windows/apps/publish/msstore-dev-cli/github-actions
- Submission resource fields (``targetPublishMode``:
  ``Immediate``/``Manual``/``SpecificDate``, ``applicationPackages``) and the
  status enum (``PendingCommit, CommitStarted, PreProcessing, Certification,
  CertificationFailed, Release, PendingPublication, Publishing, Published, ...``):
  https://learn.microsoft.com/en-us/windows/uwp/monetize/manage-app-submissions
- REST methods (get an app, get/update/commit a submission, get its status):
  https://learn.microsoft.com/en-us/windows/uwp/monetize/get-an-app
  https://learn.microsoft.com/en-us/windows/uwp/monetize/update-an-app-submission
  https://learn.microsoft.com/en-us/windows/uwp/monetize/commit-an-app-submission
  https://learn.microsoft.com/en-us/windows/uwp/monetize/get-status-for-an-app-submission
  Update and commit refuse a submission that is no longer a draft (409).
- Azure AD client-credentials token:
  https://learn.microsoft.com/en-us/windows/uwp/monetize/create-and-manage-submissions-using-windows-store-services#obtain-an-azure-ad-access-token
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request

API_ROOT = "https://manage.devcenter.microsoft.com/v1.0/my"
TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
TOKEN_SCOPE = "https://manage.devcenter.microsoft.com/.default"

_DRAFT = "PendingCommit"
_FAILED = {"CommitFailed", "PreProcessingFailed", "CertificationFailed", "PublishFailed",
           "ReleaseFailed", "Canceled"}
# Nothing pending a green run could clobber: the last submission is live or gone.
_SETTLED = {"None", "Published"}
# Committed and not live yet: certification, or certified and held.
_COMMITTED = {"CommitStarted", "PreProcessing", "Certification", "Release"}
_ALREADY_LIVE = {"PendingPublication", "Publishing", "Published"}


class StoreError(RuntimeError):
    """A Store CLI or API call failed. The calling run stays red."""


def _cli_run(argv: list[str]) -> str:
    result = subprocess.run(argv, capture_output=True, text=True, encoding="utf-8", errors="replace", check=False)
    if result.returncode != 0:
        raise StoreError(f"{' '.join(argv)} failed: {result.stderr.strip()}")
    return result.stdout


def _status(submission_output: str) -> str:
    """The status of the submission ``msstore submission get`` prints.

    ``submission status`` is no use here: for a packaged (MSIX) app it prints a Spectre
    markup line (``Submission Status = ...``) and status tables, never JSON, and prints
    nothing at all when the app has no pending submission (msstore-cli
    Commands/Submission/StatusCommand.cs). ``submission get`` serializes the submission
    as JSON (GetCommand.cs), and it is what ``submit`` already reads below.
    """
    try:
        submission = json.loads(submission_output)
    except json.JSONDecodeError:
        raise StoreError(f"unreadable submission: {submission_output[:200]!r}")
    key = _key(submission, "status") if isinstance(submission, dict) else None
    if key is None:
        raise StoreError(f"unreadable submission: {submission_output[:200]!r}")
    return submission[key]


def _key(submission: dict, name: str) -> str | None:
    """The submission's own spelling of ``name``. The CLI serializes .NET property names
    (``Status``, ``TargetPublishMode``); the REST API documents camelCase."""
    return next((k for k in submission if k.lower() == name.lower()), None)


def submit(package: str, *, product_id: str, run=_cli_run) -> dict:
    """Upload ``package`` as an uncommitted draft; the publication pass commits it."""
    status = _status(run(["msstore", "submission", "get", product_id]))
    if status == _DRAFT or status in _FAILED:
        run(["msstore", "submission", "delete", product_id, "--no-confirm"])
    elif status not in _SETTLED:
        # Committed and on its way live: an earlier release's Store update, which
        # msstore publish would delete.
        raise StoreError(f"a committed Store submission is {status}; rerun stable-store "
                         "after it is published")
    run(["msstore", "publish", package, "--appId", product_id, "--noCommit"])
    return {"productId": product_id, "status": _DRAFT}


def _http_run(request: dict) -> dict:
    """Default injected runner: one HTTP request, stdlib only."""
    url = request["url"]
    data = None
    headers = {"Content-Type": "application/json", **request.get("headers", {})}
    body = request.get("body")
    if body is not None:
        data = json.dumps(body).encode()
    if "form" in request:
        data = urllib.parse.urlencode(request["form"]).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    req = urllib.request.Request(url, data=data, headers=headers,
                                 method=request["method"])
    try:
        with urllib.request.urlopen(req) as response:
            raw = response.read().decode()
            return {"status": response.status,
                    "body": json.loads(raw) if raw else ""}
    except urllib.error.HTTPError as error:
        raw = error.read().decode()
        try:
            parsed = json.loads(raw) if raw else {}
        except json.JSONDecodeError:
            parsed = {"message": raw}
        return {"status": error.code, "body": parsed}


def _token(tenant_id: str, client_id: str, client_secret: str, run) -> str:
    response = run({"method": "POST", "url": TOKEN_URL.format(tenant=tenant_id),
                    "form": {
                        "grant_type": "client_credentials",
                        "client_id": client_id,
                        "client_secret": client_secret,
                        "scope": TOKEN_SCOPE,
                    }})
    if response["status"] != 200 or "access_token" not in response["body"]:
        raise StoreError(f"token request failed: {response['status']}")
    return response["body"]["access_token"]


def release(*, product_id: str, store_file: str, tenant_id: str, client_id: str,
            client_secret: str, run=_http_run) -> str:
    """Commit this release's Store draft with ``targetPublishMode: "Immediate"``.

    ``store_file`` is the Store bundle the release's candidate manifest names; a
    draft that does not carry it belongs to another green run and is left alone.
    """
    token = _token(tenant_id, client_id, client_secret, run)
    auth = {"Authorization": f"Bearer {token}"}

    def call(method: str, path: str, body=None) -> dict:
        request = {"method": method, "url": API_ROOT + path, "headers": auth}
        if body is not None:
            request["body"] = body
        response = run(request)
        if response["status"] != 200:
            raise StoreError(f"{method} {path} failed: {response['status']} {response['body']}")
        return response["body"]

    app = call("GET", f"/applications/{product_id}")
    pending = (app.get("pendingApplicationSubmission") or {}).get("id")
    if not pending:
        _notice(f"No Store draft is pending for {product_id}; check the release's stable-store job.")
        return "no-draft"
    path = f"/applications/{product_id}/submissions/{pending}"
    submission = call("GET", path)
    state = submission.get(_key(submission, "status") or "status")
    packages = submission.get(_key(submission, "applicationPackages") or "applicationPackages") or []
    bound = any(package.get("fileName") == store_file and package.get("fileStatus") != "PendingDelete"
                for package in packages)
    if state in _ALREADY_LIVE:
        return "already-live"
    if state in _COMMITTED:
        return "already-committed" if bound else "other-committed"
    if state in _FAILED:
        raise StoreError(f"Store submission {pending} is {state}; resubmit from a green run")
    if state != _DRAFT:
        raise StoreError(f"Store submission {pending} has an unexpected status: {state!r}")
    if not bound:
        _notice(f"The pending Store draft {pending} does not carry {store_file}; a newer "
                "green run replaced it, and that release's publication commits it.")
        return "superseded"
    submission[_key(submission, "targetPublishMode") or "targetPublishMode"] = "Immediate"
    call("PUT", path, submission)
    call("POST", f"{path}/commit")
    after = call("GET", f"{path}/status").get("status")
    if after in _FAILED:
        raise StoreError(f"Store submission {pending} is {after} after commit")
    return "committed"


def _notice(text: str) -> None:
    # A GitHub annotation, so the manual step shows on the run's summary page.
    print(f"::warning title=Microsoft Store::{text}")


def store_bundle(manifest: dict) -> str:
    """The Store bundle a candidate manifest names (the one ``materialize --store-only`` picks)."""
    names = [item["path"] for item in manifest.get("files", [])
             if item["path"].startswith("Store-") and item["path"].endswith(".msixbundle")]
    if len(names) != 1:
        raise StoreError(f"the candidate manifest names {len(names)} Store bundles, not one")
    return names[0]


def release_from_env(env: dict, manifest: dict, run=_http_run) -> str:
    """Release the Store draft for ``manifest`` when the environment configures the Store."""
    product_id = env.get("MS_STORE_PRODUCT_ID")
    if not product_id:
        print("Store release skipped: MS_STORE_PRODUCT_ID is not configured.",
              file=sys.stderr)
        return "not-configured"
    return release(
        product_id=product_id,
        store_file=store_bundle(manifest),
        tenant_id=env["MS_STORE_TENANT_ID"],
        client_id=env["MS_STORE_CLIENT_ID"],
        client_secret=env["MS_STORE_CLIENT_SECRET"],
        run=run,
    )


def main(argv: list[str]) -> int:
    if len(argv) >= 2 and argv[0] == "submit":
        product_id = os.environ["MS_STORE_PRODUCT_ID"]
        package = argv[1]
        result = submit(package, product_id=product_id)
        print(json.dumps(result, sort_keys=True))
        return 0
    print("usage: python -m scripts.releases.store submit <package>", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
