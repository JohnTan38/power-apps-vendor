# C:\apps\esker\listener.py
import os
import sys
import time
import json
import re
import shutil
import tempfile
import threading
import subprocess
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pythoncom
import win32com.client

# --- CONFIG ---
PYTHON_EXE = sys.executable  # uses current venv/interpreter
APP_PY     = r"C:/Users/john.tan/esker/Scripts/app.py" # <-- adjust if needed
APP_UI     = Path(r"C:/Users/john.tan/esker/Scripts/app_ui.py")
QUEUE_DIR = Path(r"C:/Users/john.tan/esker/queue")
ARCHIVE_SUCCESS_DIR = Path(r"C:/Users/john.tan/esker/archive/success")
KEYWORDS   = ["esker vendor email", "esker gl email"]         # lower-case match
OUTLOOK_TYPELIB_GUID = "{00062FFF-0000-0000-C000-000000000046}"
# Try known Outlook library versions (Office 2016+ typically 9.6, but keep extras for upgrades)
OUTLOOK_TYPELIB_VERSIONS = [
    (9, 8),
    (9, 7),
    (9, 6),
    (9, 5),
    (9, 4),
]
OUTLOOK_MAKEPY_SPECS = [
    "Outlook.Application",
    "Microsoft Outlook 16.0 Object Library",
]

_executor: ThreadPoolExecutor | None = None
_executor_lock = threading.Lock()

def subject_hit(subj: str) -> bool:
    s = (subj or "").lower()
    return any(k in s for k in KEYWORDS)

def write_temp_json(data: dict) -> Path:
    # Ensure queue exists
    QUEUE_DIR.mkdir(parents=True, exist_ok=True)
    # Monotonic, sortable name: YYYYmmdd_HHMMSS_mmmmmm + ns tail
    ts = time.strftime("%Y%m%d_%H%M%S")
    fname = f"{ts}_{time.time_ns()%1_000_000_000:09d}.json"
    p = QUEUE_DIR / fname
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    # Optional: keep a backup copy to Downloads as before
    downloads = Path(r"C:/Users/john.tan/Downloads/")
    downloads.mkdir(parents=True, exist_ok=True)
    shutil.copy2(p, downloads / p.name)
    return p


def ensure_outlook_typelib(app=None, retries: int = 2) -> None:
    """Ensure the Outlook type library is generated so events can attach."""
    from win32com.client import gencache, makepy

    gencache.is_readonly = False
    guid = OUTLOOK_TYPELIB_GUID
    lcid = 0
    candidates = list(OUTLOOK_TYPELIB_VERSIONS)
    makepy_specs = list(OUTLOOK_MAKEPY_SPECS)
    last_err: Exception | None = None

    if app is not None:
        try:
            type_info = app._oleobj_.GetTypeInfo()
            typelib, _ = type_info.GetContainingTypeLib()
            doc = typelib.GetDocumentation(-1)
            for spec in (doc[0], doc[1]):
                if spec and spec not in makepy_specs:
                    makepy_specs.insert(0, spec)
            lib_attr = typelib.GetLibAttr()
            guid = str(lib_attr[0])
            lcid = lib_attr[1] or 0
            dynamic = (lib_attr[3], lib_attr[4])
            if dynamic not in candidates:
                candidates.insert(0, dynamic)
        except Exception as info_err:
            _log_com_message(f"Failed to inspect Outlook typelib: {info_err}")

    attempts = max(retries, 0) + 1

    for attempt in range(attempts):
        for major, minor in candidates:
            module, err = _load_outlook_module(guid, lcid, major, minor)
            if module and _module_has_required_attrs(module):
                return

            if err is None and module:
                err = AttributeError("Outlook makepy module missing required attributes")

            if err:
                last_err = err
                _invalidate_outlook_typelib_module(guid, lcid, major, minor)
                if _generate_outlook_typelib(guid, lcid, major, minor, makepy_specs, makepy):
                    module, err = _load_outlook_module(guid, lcid, major, minor)
                    if module and _module_has_required_attrs(module):
                        return
                    if err is None and module:
                        err = AttributeError(
                            "Outlook makepy module missing required attributes after regeneration"
                        )
                if err:
                    last_err = err

        if attempt < attempts - 1:
            _clear_outlook_genpy_cache(guid)
            try:
                gencache.Rebuild()
            except Exception as rebuild_err:
                last_err = rebuild_err

    if last_err:
        raise last_err
    raise RuntimeError("Unable to prepare Outlook COM support.")


def _load_outlook_module(guid: str, lcid: int, major: int, minor: int):
    """Attempt to load the cached Outlook makepy module."""
    from win32com.client import gencache

    try:
        module = gencache.EnsureModule(guid, lcid, major, minor)
        return module, None
    except Exception as err:
        return None, err


def _module_has_required_attrs(module) -> bool:
    """Return True if the generated module exposes the attributes makepy needs."""
    required = ("CLSIDToClassMap", "CLSIDToPackageMap", "MinorVersion")
    return all(hasattr(module, attr) for attr in required)


def _invalidate_outlook_typelib_module(guid: str, lcid: int, major: int, minor: int) -> None:
    """Remove cached Outlook modules so makepy can rebuild cleanly."""
    from win32com.client import gencache

    try:
        identifier = gencache.GetGeneratedFileName(guid, lcid, major, minor)
    except Exception:
        identifier = None

    gen_dir = None
    try:
        gen_dir = Path(gencache.GetGeneratePath())
    except Exception:
        pass

    if identifier:
        module_name = f"win32com.gen_py.{identifier}"
        sys.modules.pop(module_name, None)
        if gen_dir:
            targets = [
                gen_dir / f"{identifier}.py",
                gen_dir / f"{identifier}.pyc",
                gen_dir / identifier,
            ]
            for target in targets:
                try:
                    if target.is_dir():
                        shutil.rmtree(target)
                    elif target.is_file():
                        target.unlink()
                except Exception:
                    pass
            pycache_dir = gen_dir / "__pycache__"
            if pycache_dir.exists():
                for compiled in pycache_dir.glob(f"{identifier}*.pyc"):
                    try:
                        compiled.unlink()
                    except Exception:
                        pass


def _generate_outlook_typelib(
    guid: str,
    lcid: int,
    major: int,
    minor: int,
    specs: list[str],
    makepy_module,
) -> bool:
    """Run makepy for Outlook using several spec patterns."""
    attempted = []
    seen: set[str | tuple] = set()

    for spec in specs:
        if spec and spec not in seen:
            attempted.append(spec)
            seen.add(spec)

    tuple_spec = (guid, lcid, major, minor)
    if tuple_spec not in seen:
        attempted.append(tuple_spec)
        seen.add(tuple_spec)

    for spec in attempted:
        try:
            makepy_module.GenerateFromTypeLibSpec(spec)
            _log_com_message(f"makepy generated Outlook typelib for spec {spec!r}")
            return True
        except Exception as err:
            _log_com_message(f"makepy failed for {spec!r}: {err}")

    return False


def _log_com_message(message: str) -> None:
    """Append diagnostic messages related to COM setup."""
    log_path = Path(tempfile.gettempdir()) / "esker_com_cleanup.log"
    try:
        with log_path.open("a", encoding="utf-8") as log_file:
            log_file.write(f"{time.ctime()} {message}\n")
    except Exception:
        pass


def _clear_outlook_genpy_cache(guid: str) -> None:
    """Remove cached win32com gen_py files for Outlook when they are corrupted."""
    cleared = set()
    try:
        from win32com.client import gencache
        import win32com
    except Exception:
        return

    gencache.is_readonly = False
    try:
        gen_dir = Path(gencache.GetGeneratePath())
    except Exception:
        fallback = getattr(gencache, "__gen_path__", None) or getattr(win32com, "__gen_path__", "")
        gen_dir = Path(fallback) if fallback else None

    if not gen_dir:
        return

    if not gen_dir.exists():
        return

    bare_guid = guid.strip("{}")
    patterns = {
        f"{bare_guid}*",
        f"{bare_guid.lower()}*",
        f"{bare_guid.upper()}*",
        f"{bare_guid.replace('-', '')}*",
        f"{bare_guid.replace('-', '').lower()}*",
        f"{bare_guid.replace('-', '').upper()}*",
    }

    for pattern in patterns:
        for entry in gen_dir.glob(pattern):
            try:
                if entry.is_dir():
                    shutil.rmtree(entry)
                else:
                    entry.unlink()
                cleared.add(str(entry))
            except Exception:
                continue

    for cache_dir in gen_dir.glob("__pycache__"):
        try:
            shutil.rmtree(cache_dir)
        except Exception:
            continue

    if cleared:
        log_path = Path(tempfile.gettempdir()) / "esker_com_cleanup.log"
        with log_path.open("a", encoding="utf-8") as log_file:
            log_file.write(f"{time.ctime()} Cleared Outlook gen_py cache:\n")
            for item in sorted(cleared):
                log_file.write(f"  {item}\n")


def get_outlook_namespace(retries: int = 1):
    """Return the Outlook MAPI namespace. Use EnsureDispatch where possible.

    On some systems win32com.gen_py caches a broken generated file without
    CLSIDToClassMap/CLSIDToPackageMap which causes AttributeError. If that happens, 
    remove the offending gen_py module and retry once.
    """
    try:
        # Prefer EnsureDispatch which generates early-binding helpers when needed
        app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        return app.GetNamespace("MAPI")
    except AttributeError as e:
        # Known failure modes: generated module missing CLSIDToClassMap/CLSIDToPackageMap
        if retries <= 0:
            # Final fallback: use basic Dispatch (late-binding, no cache)
            try:
                app = win32com.client.Dispatch("Outlook.Application")
                return app.GetNamespace("MAPI")
            except Exception:
                raise

        # Log the error and attempt cleanup
        err_msg = f"AttributeError in COM initialization: {e}"
        err_path = Path(tempfile.gettempdir()) / "esker_com_cleanup.log"
        with err_path.open("a", encoding="utf-8") as f:
            f.write(f"{time.ctime()} {err_msg}\n")

        # Try to identify and remove the bad gen_py module
        try:
            # Clear the gencache entirely - more aggressive approach
            try:
                win32com.client.gencache.is_readonly = False
                # Try to rebuild the entire cache
                win32com.client.gencache.Rebuild()
            except Exception as rebuild_err:
                # If rebuild fails, manually remove files
                try:
                    import importlib
                    gen_py = importlib.import_module("win32com.gen_py")
                    # Remove all Outlook typelib variants
                    for p in gen_py.__path__:
                        pth = Path(p)
                        removed_files = []
                        for pattern in ['00062FFF-0000-0000-C000-000000000046*', '*Outlook*']:
                            for f in pth.glob(pattern):
                                try:
                                    if f.is_file():
                                        f.unlink()
                                        removed_files.append(str(f))
                                except Exception:
                                    pass
                        # Also remove __pycache__ folders that might contain stale bytecode
                        for cache_dir in pth.glob('__pycache__'):
                            try:
                                if cache_dir.is_dir():
                                    shutil.rmtree(cache_dir)
                            except Exception:
                                pass
                        
                        with err_path.open("a", encoding="utf-8") as f:
                            f.write(f"{time.ctime()} Removed files: {removed_files}\n")
                except Exception as cleanup_err:
                    with err_path.open("a", encoding="utf-8") as f:
                        f.write(f"{time.ctime()} Cleanup error: {cleanup_err}\n")
        except Exception:
            # If we cannot clean, re-raise original error
            raise

        # Retry once
        return get_outlook_namespace(retries=retries - 1)

class InboxEvents:
    # IMPORTANT: no __init__(self, items) here! WithEvents calls with no args.
    def OnItemAdd(self, item):
        try:
            # 43 = olMailItem
            if getattr(item, "Class", None) != 43:
                return

            subj = getattr(item, "Subject", "")
            if not subject_hit(subj):
                return

            sender = getattr(item, "SenderEmailAddress", "")
            body   = getattr(item, "Body", "")
            recv   = item.ReceivedTime  # COM Date
            try:
                recv_str = recv.strftime("%Y-%m-%dT%H:%M:%SZ")
            except Exception:
                recv_str = str(recv)

            payload = {
                "subject": subj,
                "sender_address": sender,
                "received_utc": recv_str,
                "body": body,
                "entry_id": getattr(item, "EntryID", ""),
            }

            json_path = write_temp_json(payload)
            # Launch your worker (non-blocking)
            # Queue worker model (Option A): do NOT launch app.py here.
            # A long-running worker (app_ui.py --mode=worker) will pick this up.
            # If you want to auto-start the worker when idle, you can add logic here.

            # Optional: quick console note
            print(f"[listener] Triggered for: {subj} -> {json_path}")

            enqueue_worker(json_path)

        except Exception as e:
            # Minimal error logging to temp
            err_path = Path(tempfile.gettempdir()) / "esker_listener_errors.log"
            with err_path.open("a", encoding="utf-8") as f:
                f.write(f"{time.ctime()} OnItemAdd error: {e}\n")


def enqueue_worker(json_path: Path) -> None:
    """Queue a JSON payload for processing by the background worker."""
    print(f"[listener] Queued automation job for {json_path.name}")
    executor = ensure_worker_executor()
    executor.submit(worker_task, json_path)


def ensure_worker_executor() -> ThreadPoolExecutor:
    """Return a singleton ThreadPoolExecutor used for automation runs."""
    global _executor
    with _executor_lock:
        if _executor is None:
            _executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="esker-worker")
        return _executor


def worker_task(json_path: Path) -> None:
    """Wrapper submitted to the executor to run the automation for a payload."""
    try:
        run_worker(json_path)
    except Exception as exc:
        log_path = Path(tempfile.gettempdir()) / "esker_listener_runner.log"
        with log_path.open("a", encoding="utf-8") as log_file:
            log_file.write(f"{time.ctime()} worker_task error for {json_path.name}: {exc}\n")


def run_worker(json_path: Path) -> None:
    """Invoke app_ui.py for the provided payload and archive on success."""
    env = os.environ.copy()
    env.pop("ESKER_DRYRUN", None)
    env["ESKER_VENDOR_JSON_DIR"] = str(json_path.parent)
    env["ESKER_VENDOR_JSON_PATTERN"] = json_path.name

    cmd = [PYTHON_EXE, str(APP_UI), "--mode=worker"]
    log_path = Path(tempfile.gettempdir()) / "esker_listener_runner.log"
    print(f"[listener] Starting automation for {json_path.name}")
    result = subprocess.run(cmd, cwd=str(APP_UI.parent), env=env, capture_output=True, text=True)

    with log_path.open("a", encoding="utf-8") as log_file:
        log_file.write(
            f"{time.ctime()} ran {cmd} for {json_path.name} -> rc={result.returncode}\n"
        )
        if result.stdout:
            log_file.write(f"stdout:\n{result.stdout}\n")
        if result.stderr:
            log_file.write(f"stderr:\n{result.stderr}\n")

    if result.returncode == 0:
        ARCHIVE_SUCCESS_DIR.mkdir(parents=True, exist_ok=True)
        shutil.copy2(json_path, ARCHIVE_SUCCESS_DIR / json_path.name)
        try:
            json_path.unlink()
        except OSError:
            pass
        print(f"[listener] Worker completed; archived to {ARCHIVE_SUCCESS_DIR / json_path.name}")
    else:
        print(f"[listener] Worker failed for {json_path.name}; see {log_path}")


def main():
    # Get Outlook namespace and Inbox folder
    outlook = get_outlook_namespace()
    inbox   = outlook.GetDefaultFolder(6)  # 6 = olFolderInbox

    # Keep strong references so events stay alive
    items = inbox.Items
    # NOTE: Do not Sort/Restrict here; ItemAdd fires only on the default Items collection.
    # If you need filtering, do it in OnItemAdd.

    # Hook events
    ensure_outlook_typelib(app=outlook.Application)
    try:
        handler = win32com.client.WithEvents(items, InboxEvents)
    except TypeError as err:
        ensure_outlook_typelib(app=outlook.Application, retries=1)
        try:
            handler = win32com.client.WithEvents(items, InboxEvents)
        except TypeError as final_err:
            raise TypeError(
                "Outlook event binding failed even after rebuilding makepy cache."
            ) from final_err

    print("Listening for new Inbox items... (Ctrl+C to exit)")
    # Pump COM messages
    while True:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.2)

if __name__ == "__main__":
    main()
