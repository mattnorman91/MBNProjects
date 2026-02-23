#!/usr/bin/env python3
# main.py
"""
MBN Projects — Backend and Launcher (robust logo handling, full file)

Features:
- Prepares logo files next to the app (logo_injected.svg or logo.png and logo_small.png).
- Reads source file bytes first (safer on Windows when file is locked), writes to target.
- Falls back to file:// URI only if read/write/copy fails.
- Handles PyInstaller --onefile via _MEIPASS.
- Writes index_injected.html (replacing {{LOGO_SRC}}) and opens it with webview.create_window(url=...).
- Full API (list_recent, new_project, open_excel, save_project, quick_summary, etc.)
- Windows host monitor: detects minimize -> restore and calls JS redraw so charts/layout repaint correctly.
"""
import sys
import os
import json
import uuid
import shutil
import time
import atexit
import signal
import logging
import re
import threading
from pathlib import Path
from datetime import datetime, date
from tkinter import Tk, filedialog
from io import BytesIO

# Optional dependencies
try:
    import pandas as pd
except Exception:
    pd = None

try:
    from PIL import Image
except Exception:
    Image = None

try:
    import cairosvg
except Exception:
    cairosvg = None

try:
    import webview
except Exception:
    webview = None

# Windows-specific: try to import win32gui for restore detection
try:
    import win32gui
except Exception:
    win32gui = None

logging.basicConfig(stream=sys.stderr, level=logging.DEBUG, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger("mbnprojects")

# ---------- base storage directories ----------
def get_base_dir():
    """
    Use the user's Documents folder as the base storage location:
      <User Documents>/MBNProjects
    Falls back to ~/.mbn_projects if Documents isn't available.
    """
    try:
        home = Path.home()
        docs = None
        if os.name == 'nt':
            docs = Path(os.path.join(os.environ.get('USERPROFILE', str(home)), 'Documents'))
        else:
            docs = home / "Documents"
        if docs and docs.exists():
            base = docs / "MBNProjects"
        else:
            base = home / "MBNProjects"
    except Exception:
        base = Path.home() / "MBNProjects"
    return base

BASE_DIR = get_base_dir()
TEMPLATES_DIR = BASE_DIR / "templates"
PROJECTS_DIR = BASE_DIR / "projects"
SETTINGS_FILE = BASE_DIR / "settings.json"

# Ensure directories exist
BASE_DIR.mkdir(parents=True, exist_ok=True)
TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
PROJECTS_DIR.mkdir(parents=True, exist_ok=True)

logger.debug("Base directory: %s", BASE_DIR)
logger.debug("Templates dir: %s", TEMPLATES_DIR)
logger.debug("Projects dir: %s", PROJECTS_DIR)
logger.debug("Settings file: %s", SETTINGS_FILE)

# ---------- utilities ----------
def atomic_write_text(path: Path, text: str):
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(str(tmp), str(path))
    logger.debug("Atomic write complete: %s", path)

def choose_file_dialog(save=False, defaultextension=".xlsx", initialdir=None):
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = None
    try:
        if save:
            path = filedialog.asksaveasfilename(defaultextension=defaultextension, filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")], initialdir=str(initialdir) if initialdir else None)
        else:
            path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm"),("All Files","*.*")], initialdir=str(initialdir) if initialdir else None)
    finally:
        try:
            root.destroy()
        except Exception:
            pass
    return path

def normalize_path_for_storage(p):
    try:
        rp = Path(p).expanduser().resolve()
        s = str(rp)
    except Exception:
        s = str(Path(p).expanduser().absolute())
    if os.name == 'nt':
        s = s.lower()
    return s

def safe_folder_name(name: str) -> str:
    if not name:
        name = "Untitled Project"
    safe = "".join(c for c in name if c.isalnum() or c in (" ", "_", "-")).rstrip()
    safe = safe.strip().replace(" ", "_")
    if not safe:
        safe = "project"
    return safe

def create_project_structure(base_projects_dir: Path, project_name: str) -> Path:
    try:
        safe = safe_folder_name(project_name)
        candidate = base_projects_dir / safe
        if candidate.exists():
            unique = int(time.time())
            candidate = base_projects_dir / f"{safe}_{unique}"
        candidate.mkdir(parents=True, exist_ok=True)
        (candidate / "Planning").mkdir(parents=True, exist_ok=True)
        (candidate / "RAID").mkdir(parents=True, exist_ok=True)
        (candidate / "Attachments").mkdir(parents=True, exist_ok=True)
        logger.info("Created project structure at %s", candidate)
        return candidate
    except Exception:
        logger.exception("Failed to create project structure")
        raise

def date_to_iso(d):
    if d is None:
        return ""
    try:
        if isinstance(d, pd.Timestamp):
            if pd.isna(d):
                return ""
            return d.date().isoformat()
    except Exception:
        pass
    try:
        if isinstance(d, datetime):
            return d.date().isoformat()
        if isinstance(d, date):
            return d.isoformat()
    except Exception:
        pass
    try:
        parsed = pd.to_datetime(d, errors='coerce')
        if not pd.isna(parsed):
            return parsed.date().isoformat()
    except Exception:
        pass
    try:
        if isinstance(d, (int, float)):
            return datetime.fromtimestamp(d).date().isoformat()
    except Exception:
        pass
    try:
        if pd is not None and pd.isna(d):
            return ""
    except Exception:
        pass
    return str(d)

def sanitize_project_dict(d: dict) -> dict:
    safe = {}
    for k, v in (d or {}).items():
        try:
            if pd is not None and isinstance(v, (pd.Timestamp, datetime, date)):
                safe[k] = date_to_iso(v)
            else:
                if pd is not None and ((isinstance(v, float) and pd.isna(v)) or pd.isna(v)):
                    safe[k] = ""
                else:
                    if isinstance(v, (bool, int)):
                        safe[k] = v
                    elif isinstance(v, list):
                        safe[k] = v
                    else:
                        safe[k] = "" if v is None else str(v)
        except Exception:
            safe[k] = "" if v is None else str(v)
    return safe

# ---------- settings helpers ----------
def is_temporary_path(p):
    try:
        name = Path(p).name.lower()
        return ('.mbnwork.' in name) or ('.autosave.' in name)
    except Exception:
        return False

def load_settings_raw():
    if not SETTINGS_FILE.exists():
        return {"recent": [], "templates": [], "ui": {}}
    try:
        raw = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            return {"recent": [], "templates": [], "ui": {}}
    except Exception:
        return {"recent": [], "templates": [], "ui": {}}

    raw_recents = raw.get("recent", []) if isinstance(raw.get("recent", []), list) else []
    seen = set()
    normalized_list = []
    for item in raw_recents:
        try:
            norm = normalize_path_for_storage(item)
        except Exception:
            norm = str(item)
        if is_temporary_path(norm):
            continue
        if norm in seen:
            continue
        seen.add(norm)
        normalized_list.append(norm)
    raw["recent"] = normalized_list
    raw.setdefault("templates", [])
    raw.setdefault("ui", {})
    try:
        atomic_write_text(SETTINGS_FILE, json.dumps(raw, indent=2))
    except Exception:
        logger.debug("Could not persist normalized settings immediately")
    return raw

def save_settings_raw(settings):
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    atomic_write_text(SETTINGS_FILE, json.dumps(settings, indent=2))

def add_recent_raw(path_str):
    if is_temporary_path(path_str):
        logger.debug("Skipping adding temporary file to recents: %s", path_str)
        return
    settings = load_settings_raw()
    recents = settings.get("recent", [])
    norm = normalize_path_for_storage(path_str)
    recents = [r for r in recents if r != norm]
    recents.insert(0, norm)
    settings["templates"] = settings.get("templates", [])
    settings["recent"] = recents[:12]
    save_settings_raw(settings)

def list_templates_raw():
    settings = load_settings_raw()
    return settings.get("templates", [])

def add_template_raw(meta):
    settings = load_settings_raw()
    templates = settings.get("templates", [])
    templates = [t for t in templates if t.get("filename") != meta.get("filename")]
    templates.insert(0, meta)
    settings["templates"] = templates
    save_settings_raw(settings)

# ---------- Column mapping & data handling ----------
COMMON_TASK_COLS = {
    "taskid":"Task ID","id":"Task ID","task id":"Task ID",
    "task_name":"Task Name","taskname":"Task Name","task name":"Task Name","name":"Task Name",
    "startdate":"Start Date","start_date":"Start Date","start date":"Start Date","start":"Start Date",
    "enddate":"End Date","end_date":"End Date","end date":"End Date","end":"End Date",
    "duration":"Duration","status":"Status","dependencies":"Dependencies","depends":"Dependencies",
    "resource":"Resource","blocking":"Blocking","block":"Blocking","lag":"Lag","order":"Order","group":"Group"
}

def _normalize_col_name(name: str) -> str:
    if not isinstance(name, str):
        return ""
    return name.strip().replace(" ", "_").replace("-", "_").replace(".", "_").lower()

def _map_columns(cols):
    mapping = {}
    normalized_to_original = { _normalize_col_name(c): c for c in cols }
    for norm, orig in normalized_to_original.items():
        if norm in COMMON_TASK_COLS:
            canonical = COMMON_TASK_COLS[norm]
            if canonical not in mapping:
                mapping[canonical] = orig
    return mapping

def _ensure_columns_and_types(df):
    expected = ['Task ID','Task Name','Start Date','End Date','Duration','Status','Dependencies','Resource','Blocking','Lag','Order','Group']
    for c in expected:
        if c not in df.columns:
            df[c] = pd.NA
    df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce')
    df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce')
    df['Duration'] = pd.to_numeric(df['Duration'], errors='coerce')
    def parse_blocking(v):
        if pd.isna(v): return False
        if isinstance(v, bool): return v
        if isinstance(v, (int,float)): return bool(v)
        sv = str(v).strip().lower()
        return sv in ('1','true','yes','y','t')
    df['Blocking'] = df['Blocking'].apply(parse_blocking)
    df['Lag'] = pd.to_numeric(df['Lag'], errors='coerce').fillna(0).astype(int)
    df['Duration'] = df['Duration'].fillna((df['End Date'] - df['Start Date']).dt.days + 1).fillna(0).astype(int)
    def ensure_id(v):
        if pd.isna(v) or v == "": return str(uuid.uuid4())
        return str(v)
    df['Task ID'] = df['Task ID'].apply(ensure_id)
    if 'Order' not in df.columns or df['Order'].isna().all():
        df['Order'] = range(1, len(df)+1)
    else:
        df['Order'] = pd.to_numeric(df['Order'], errors='coerce').fillna(0).astype(int)
    df['Dependencies'] = df['Dependencies'].fillna("").astype(str)
    df['Resource'] = df['Resource'].fillna("").astype(str)
    df['Group'] = df['Group'].fillna("").astype(str)
    return df[expected]

# ---------- API ----------
class Api:
    def __init__(self):
        self.last_path = None
        self.original_path = None
        self.working_path = None
        self._save_lock = False
        try:
            atexit.register(self._atexit_cleanup)
        except Exception:
            logger.exception("Failed to register atexit cleanup")
        try:
            signal.signal(signal.SIGINT, self._signal_handler)
            try:
                signal.signal(signal.SIGTERM, self._signal_handler)
            except Exception:
                pass
        except Exception:
            pass

    def _atexit_cleanup(self):
        try:
            if self.working_path:
                wp = Path(self.working_path)
                if wp.exists():
                    try:
                        wp.unlink(missing_ok=True)
                        logger.debug("Removed working copy during atexit: %s", wp)
                    except Exception:
                        logger.exception("Failed to remove working copy during atexit: %s", wp)
                self.working_path = None
        except Exception:
            logger.exception("Error during atexit cleanup")

    def _signal_handler(self, signum, frame):
        logger.debug("Signal %s received; running cleanup", signum)
        try:
            self._atexit_cleanup()
        finally:
            try:
                sys.exit(0)
            except Exception:
                pass

    def _cleanup_temp_files_for_original(self, orig: Path):
        try:
            parent = orig.parent
            stem = orig.stem.lower()
            for child in parent.iterdir():
                try:
                    name = child.name.lower()
                except Exception:
                    logger.exception("Failed reading child name for %s", child)
                    continue
                if stem in name and ('.mbnwork.' in name or '.autosave.' in name):
                    try:
                        child.unlink(missing_ok=True)
                        logger.debug("Removed temp file: %s", child)
                    except Exception:
                        logger.exception("Failed to remove temp file %s", child)
        except Exception:
            logger.exception("cleanup temp files failed")

    # ---------- public API methods ----------
    def get_current_path(self):
        return {"path": self.last_path or ""}

    def get_settings_path(self):
        return {"settings_path": str(SETTINGS_FILE)}

    def new_project(self, project_name: str = None):
        """
        Create a new project with the provided project_name.
        Called from JS as: await window.pywebview.api.new_project(projectName)
        Returns {"project": {...}, "tasks": [], "project_folder": "<path>"} or {"error": "msg"}
        """
        try:
            if not project_name or not isinstance(project_name, str):
                return {"error": "Missing project name"}
            name = project_name.strip()
            if not name:
                return {"error": "Missing project name"}

            logger.debug("new_project called with name: %s", name)
            # Create directory structure
            folder = create_project_structure(PROJECTS_DIR, name)

            # Ensure Planning folder exists (create_project_structure already creates it)
            planning = folder / "Planning"
            planning.mkdir(parents=True, exist_ok=True)

            # Create a Planning workbook (plan.xlsx) in the Planning folder.
            plan_xlsx = planning / "plan.xlsx"
            plan_csv = planning / "plan.csv"
            try:
                if pd is not None:
                    try:
                        proj_df = pd.DataFrame([{
                            "Project Name": name,
                            "Created": datetime.now().isoformat(),
                            "Owner": "",
                            "Notes": "",
                            "Groups": []
                        }])
                        tasks_df = pd.DataFrame(columns=['Task ID','Task Name','Start Date','End Date','Duration','Status','Dependencies','Resource','Blocking','Lag','Order','Group'])
                        with pd.ExcelWriter(str(plan_xlsx), engine='openpyxl') as writer:
                            proj_df.to_excel(writer, sheet_name='Project', index=False)
                            tasks_df.to_excel(writer, sheet_name='Tasks', index=False)
                        logger.info("Created planning workbook at %s", plan_xlsx)
                    except Exception:
                        logger.exception("Failed creating plan.xlsx via pandas; falling back to CSV + placeholder")
                        plan_csv.write_text("Task,Owner,Start,End,Status\n", encoding="utf-8")
                        try:
                            plan_xlsx.write_bytes(b"")
                        except Exception:
                            pass
                else:
                    plan_csv.write_text("Task,Owner,Start,End,Status\n", encoding="utf-8")
                    try:
                        plan_xlsx.write_bytes(b"")
                    except Exception:
                        pass
            except Exception:
                logger.exception("Failed creating planning files")
                return {"error": "Failed creating planning files"}

            # Update recent and state
            self.original_path = None
            self.working_path = None
            self.last_path = str(planning)
            add_recent_raw(str(planning))
            logger.info("New project created with folder: %s", folder)
            project = {
                "Project Name": name,
                "Created": datetime.now().isoformat(),
                "Owner": "",
                "Notes": "",
                "Groups": [],
                "Project Folder": str(folder)
            }
            return {"project": project, "tasks": [], "project_folder": str(folder)}
        except Exception:
            logger.exception("new_project failed")
            return {"error": "new_project failed"}

    def create_project_dir(self, project_name):
        try:
            folder = create_project_structure(PROJECTS_DIR, project_name)
            return {"ok": True, "project_folder": str(folder)}
        except Exception:
            logger.exception("create_project_dir failed")
            return {"error": "create_project_dir failed"}

    def list_recent(self):
        settings = load_settings_raw()
        recents = settings.get("recent", [])
        out = []
        for p in recents:
            if is_temporary_path(p):
                continue
            exists = False
            display = p
            last_modified = ""
            try:
                candidate = Path(p)
                if candidate.exists():
                    display = str(candidate.resolve()); exists = True
                    try:
                        last_modified = datetime.fromtimestamp(candidate.stat().st_mtime).isoformat()
                    except Exception:
                        last_modified = ""
                else:
                    if os.name == 'nt':
                        parent = candidate.parent; name = candidate.name
                        if parent.exists():
                            for child in parent.iterdir():
                                if child.name.lower() == name.lower():
                                    display = str(child.resolve()); exists = True
                                    try: last_modified = datetime.fromtimestamp(child.stat().st_mtime).isoformat()
                                    except Exception: last_modified = ""
                                    break
            except Exception:
                pass
            out.append({"path": display, "exists": bool(exists), "last_modified": last_modified})
        return {"recent": out}

    def list_templates(self):
        settings = load_settings_raw()
        return {"templates": settings.get("templates", [])}

    def upload_template(self):
        path = choose_file_dialog(save=False)
        if not path:
            return {"error": "No file selected"}
        try:
            src = Path(path)
            target_name = f"{src.stem}_{int(datetime.now().timestamp())}{src.suffix}"
            target = TEMPLATES_DIR / target_name
            tmp_target = target.with_suffix(target.suffix + ".tmp")
            shutil.copy2(src, tmp_target)
            os.replace(str(tmp_target), str(target))
            meta = {"name": src.stem, "filename": target_name, "uploaded": datetime.now().isoformat()}
            add_template_raw(meta)
            return {"ok": True, "template": meta}
        except Exception:
            logger.exception("upload_template failed")
            return {"error": "upload_template failed"}

    def _create_working_copy(self, original_path: str) -> str:
        try:
            orig = Path(original_path)
            parent = orig.parent if orig.parent.exists() else Path.cwd()
            unique = uuid.uuid4().hex
            working_name = f"{orig.stem}.mbnwork.{unique}{orig.suffix}"
            working = parent / working_name
            if orig.exists():
                try:
                    shutil.copy2(orig, working)
                except PermissionError:
                    logger.warning("PermissionError copying original to working copy; creating empty workbook instead")
                    if pd:
                        empty_proj = pd.DataFrame([{"Project Name": "", "Created": datetime.now().isoformat(), "Owner": "", "Notes": "", "Groups": "[]"}])
                        empty_tasks = pd.DataFrame(columns=['Task ID','Task Name','Start Date','End Date','Duration','Status','Dependencies','Resource','Blocking','Lag','Order','Group'])
                        with pd.ExcelWriter(str(working), engine='openpyxl') as writer:
                            empty_proj.to_excel(writer, sheet_name='Project', index=False)
                            empty_tasks.to_excel(writer, sheet_name='Tasks', index=False)
                except Exception:
                    logger.exception("Failed to copy original to working copy; creating empty workbook")
                    if pd:
                        empty_proj = pd.DataFrame([{"Project Name": "", "Created": datetime.now().isoformat(), "Owner": "", "Notes": "", "Groups": "[]"}])
                        empty_tasks = pd.DataFrame(columns=['Task ID','Task Name','Start Date','End Date','Duration','Status','Dependencies','Resource','Blocking','Lag','Order','Group'])
                        with pd.ExcelWriter(str(working), engine='openpyxl') as writer:
                            empty_proj.to_excel(writer, sheet_name='Project', index=False)
                            empty_tasks.to_excel(writer, sheet_name='Tasks', index=False)
            else:
                if pd:
                    empty_proj = pd.DataFrame([{"Project Name": "", "Created": datetime.now().isoformat(), "Owner": "", "Notes": "", "Groups": "[]"}])
                    empty_tasks = pd.DataFrame(columns=['Task ID','Task Name','Start Date','End Date','Duration','Status','Dependencies','Resource','Blocking','Lag','Order','Group'])
                    with pd.ExcelWriter(str(working), engine='openpyxl') as writer:
                        empty_proj.to_excel(writer, sheet_name='Project', index=False)
                        empty_tasks.to_excel(writer, sheet_name='Tasks', index=False)
            logger.debug("Created working copy: %s", working)
            return str(working)
        except Exception:
            logger.exception("Failed to create working copy")
            return original_path

    def _open_file_path(self, path):
        try:
            if pd is None:
                logger.warning("pandas not installed - returning minimal project info")
                self.original_path = str(path)
                self.working_path = None
                self.last_path = str(path)
                return {"project": {"Project Name": Path(path).stem}, "tasks": [], "path": path}

            xls = pd.read_excel(path, sheet_name=None, engine='openpyxl')
            project = {}
            tasks = []
            if 'Project' in xls:
                proj_df = xls['Project']
                if not proj_df.empty:
                    project = proj_df.iloc[0].to_dict()
                    project = sanitize_project_dict(project)
            tasks_df = None
            if 'Tasks' in xls:
                tasks_df = xls['Tasks']
            else:
                for sheet_name, df in xls.items():
                    cols = list(df.columns)
                    mapping = _map_columns(cols)
                    if 'Task Name' in mapping or 'Start Date' in mapping or 'Status' in mapping:
                        tasks_df = df
                        break
            if tasks_df is None:
                self.original_path = str(path)
                try:
                    self.working_path = self._create_working_copy(self.original_path)
                except Exception:
                    self.working_path = None
                self.last_path = str(path)
                return {"project": project or {"Project Name":"Imported Project","Created":datetime.now().isoformat(),"Groups":[]}, "tasks": [], "path": path}

            raw_cols = list(tasks_df.columns)
            mapping = _map_columns(raw_cols)
            rename_map = { actual:canon for canon, actual in mapping.items() }
            if rename_map:
                tasks_df = tasks_df.rename(columns=rename_map)
            tasks_df = _ensure_columns_and_types(tasks_df)
            tasks_df = tasks_df.sort_values('Order').reset_index(drop=True)
            out_tasks = []
            for _, row in tasks_df.iterrows():
                out_tasks.append({
                    "Task ID": str(row['Task ID']),
                    "Task Name": "" if pd.isna(row['Task Name']) else str(row['Task Name']),
                    "Start Date": date_to_iso(row['Start Date']),
                    "End Date": date_to_iso(row['End Date']),
                    "Duration": int(row['Duration']) if not pd.isna(row['Duration']) else 0,
                    "Status": "" if pd.isna(row['Status']) else str(row['Status']),
                    "Dependencies": "" if pd.isna(row['Dependencies']) else str(row['Dependencies']),
                    "Resource": "" if pd.isna(row['Resource']) else str(row['Resource']),
                    "_blocking": bool(row['Blocking']),
                    "_lag": int(row['Lag']) if not pd.isna(row['Lag']) else 0,
                    "Order": int(row['Order']) if not pd.isna(row['Order']) else None,
                    "Group": "" if pd.isna(row.get('Group', "")) else str(row.get('Group', ""))
                })
            self.original_path = str(path)
            try:
                self.working_path = self._create_working_copy(self.original_path)
            except Exception:
                self.working_path = None
            self.last_path = str(path)
            return {"project": project or {"Project Name":"Imported Project","Created":datetime.now().isoformat(),"Groups":[]}, "tasks": out_tasks, "path": path}
        except Exception:
            logger.exception("_open_file_path failed")
            return {"error": "Failed to open file; check console for details"}

    def open_path(self, path):
        if not path:
            return {"error": "No path provided"}
        p = Path(path)
        if not p.exists():
            return {"error": "File not found"}
        res = self._open_file_path(str(p))
        if "error" not in res:
            add_recent_raw(str(p))
        return res

    def open_excel(self):
        path = choose_file_dialog(save=False)
        if not path:
            return {"error": "No file selected"}
        res = self._open_file_path(path)
        if "error" not in res:
            add_recent_raw(path)
        return res

    def _commit_working_to_original(self, working: str, original: str):
        try:
            w = Path(working)
            orig = Path(original)
            if not w.exists():
                return {"error": "Working copy missing; save failed"}
            for attempt in range(4):
                try:
                    os.replace(str(w), str(orig))
                    logger.info("Committed working copy to original: %s -> %s", w, orig)
                    add_recent_raw(str(orig))
                    self.last_path = str(orig)
                    self.working_path = None
                    self.original_path = str(orig)
                    try:
                        self._cleanup_temp_files_for_original(orig)
                    except Exception:
                        logger.exception("cleanup after commit failed")
                    return {"ok": True, "path": str(orig)}
                except PermissionError as pe:
                    logger.warning("PermissionError committing working -> original (attempt %d): %s", attempt+1, pe)
                    time.sleep(0.35)
                except Exception:
                    logger.exception("Failed to replace working with original")
                    break

            try:
                ts = int(time.time())
                autosave_name = f"{orig.stem}.autosave.{ts}{orig.suffix}"
                autosave_path = orig.with_name(autosave_name)
                os.replace(str(w), str(autosave_path))
                self.last_path = str(autosave_path)
                self.working_path = None
                logger.info("Wrote autosave to %s because original was locked", autosave_path)
                return {"ok": True, "path": str(autosave_path), "warning": "Original file locked; saved to autosave file."}
            except Exception:
                logger.exception("Failed to move working copy to autosave fallback")
                return {"error": "Permission denied and autosave failed; close the original file and try again."}
        except Exception:
            logger.exception("commit_working_to_original failed")
            return {"error": "Failed to commit working copy"}

    def _write_project_file(self, project, tasks_payload, target_path):
        try:
            if pd is None:
                tgt = Path(target_path)
                tgt.parent.mkdir(parents=True, exist_ok=True)
                data = {"project": project, "tasks": tasks_payload}
                with open(tgt.with_suffix('.json'), 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2)
                add_recent_raw(str(tgt.with_suffix('.json')))
                self.last_path = str(tgt.with_suffix('.json'))
                return {"ok": True, "path": str(tgt.with_suffix('.json')), "note": "pandas not available; wrote JSON fallback"}
            rows = []
            for idx, t in enumerate(tasks_payload):
                deps = t.get("Dependencies", "") if "Dependencies" in t else ""
                if isinstance(deps, list):
                    deps = ", ".join(deps)
                blocking = t.get("_blocking", t.get("Blocking", False))
                lag = t.get("_lag", t.get("Lag", 0))
                task_id = t.get("Task ID") or str(uuid.uuid4())
                start = t.get("Start Date", "")
                end = t.get("End Date", "")
                duration = t.get("Duration", None)
                try:
                    start_val = pd.to_datetime(start, errors='coerce') if start else pd.NaT
                    end_val = pd.to_datetime(end, errors='coerce') if end else pd.NaT
                except:
                    start_val = pd.NaT; end_val = pd.NaT
                rows.append({
                    "Task ID": task_id,
                    "Task Name": t.get("Task Name", ""),
                    "Start Date": start_val,
                    "End Date": end_val,
                    "Duration": duration if duration is not None else pd.NA,
                    "Status": t.get("Status", ""),
                    "Dependencies": deps,
                    "Resource": t.get("Resource", ""),
                    "Blocking": bool(blocking),
                    "Lag": int(lag) if lag is not None else 0,
                    "Order": idx+1,
                    "Group": t.get("Group", "") or ""
                })

            tasks_df = pd.DataFrame(rows)
            tasks_df = _ensure_columns_and_types(tasks_df)

            groups_json = json.dumps(project.get("Groups", []))
            proj_row = {"Project Name": project.get("Project Name",""), "Created": project.get("Created", datetime.now().isoformat()), "Owner": project.get("Owner",""), "Notes": project.get("Notes",""), "Groups": groups_json}
            proj_df = pd.DataFrame([proj_row])

            target = Path(target_path)
            target.parent.mkdir(parents=True, exist_ok=True)

            unique = uuid.uuid4().hex
            tmp_file = target.with_name(f"{target.stem}.{unique}.tmp{target.suffix}")

            with pd.ExcelWriter(str(tmp_file), engine='openpyxl') as writer:
                proj_df.to_excel(writer, sheet_name='Project', index=False)
                tasks_df.to_excel(writer, sheet_name='Tasks', index=False)

            try:
                os.replace(str(tmp_file), str(target))
            except PermissionError as pe:
                logger.warning("PermissionError replacing tmp -> target, attempting autosave fallback: %s", pe)
                try:
                    ts = int(time.time())
                    autosave_name = f"{target.stem}.autosave.{ts}{target.suffix}"
                    autosave_path = target.with_name(autosave_name)
                    os.replace(str(tmp_file), str(autosave_path))
                    self.last_path = str(autosave_path)
                    logger.info("Wrote autosave to %s because original was locked", autosave_path)
                    return {"ok": True, "path": str(autosave_path), "warning": "Original file locked; saved to autosave file."}
                except Exception:
                    logger.exception("Failed to write autosave fallback")
                    try:
                        if tmp_file.exists():
                            tmp_file.unlink(missing_ok=True)
                    except Exception:
                        pass
                    return {"error": "Permission denied and autosave failed; close the file and try again."}
            except Exception:
                logger.exception("Failed to replace temporary file with target")
                try:
                    if tmp_file.exists():
                        tmp_file.unlink(missing_ok=True)
                except Exception:
                    pass
                return {"error": "Failed to write project file"}

            if not is_temporary_path(str(target)):
                add_recent_raw(str(target))
            self.last_path = str(target)
            logger.debug("Wrote project to %s", target)
            return {"ok": True, "path": str(target)}
        except Exception:
            logger.exception("_write_project_file failed")
            return {"error": "Failed to write project file"}

    def save_project(self, payload):
        if self._save_lock:
            logger.debug("Save requested but another save is in progress")
            return {"error": "Save already in progress"}
        self._save_lock = True
        try:
            project = payload.get("project", {})
            tasks_payload = payload.get("tasks", [])
            # If we have a working copy, write there first.
            if self.working_path:
                write_res = self._write_project_file(project, tasks_payload, self.working_path)
                if write_res.get("error"):
                    return write_res
                if self.original_path:
                    commit_res = self._commit_working_to_original(self.working_path, self.original_path)
                    return commit_res
                return write_res

            # If last_path is a folder (e.g., Planning folder) we should default to a file name inside it.
            if self.last_path:
                lp = Path(self.last_path)
                target_path = None
                if lp.is_dir():
                    fname = f"{project.get('Project Name','Project')}.xlsx"
                    fname = safe_folder_name(fname)
                    target_path = str(lp / fname)
                else:
                    target_path = str(lp)
                res = self._write_project_file(project, tasks_payload, target_path)
                try:
                    if res.get("ok") and self.original_path:
                        lastp = Path(self.last_path)
                        origp = Path(self.original_path)
                        try:
                            if lastp.exists() and origp.exists() and lastp.samefile(origp):
                                self._cleanup_temp_files_for_original(origp)
                        except Exception:
                            logger.exception("cleanup after direct write failed")
                except Exception:
                    logger.exception("cleanup after direct write failed")
                return res
            return self.save_project_as(payload)
        except Exception:
            logger.exception("save_project failed")
            return {"error": "save_project failed"}
        finally:
            self._save_lock = False

    def save_project_as(self, payload):
        try:
            project = payload.get("project", {})
            tasks_payload = payload.get("tasks", [])
            initialdir = None
            proj_folder = project.get("Project Folder") or project.get("project_folder") or None
            if proj_folder:
                try:
                    planning = Path(proj_folder) / "Planning"
                    if planning.exists():
                        initialdir = planning
                except Exception:
                    initialdir = None
            if not initialdir:
                initialdir = PROJECTS_DIR

            target_path = choose_file_dialog(save=True, initialdir=initialdir)
            if not target_path:
                return {"error": "Save cancelled"}
            res = self._write_project_file(project, tasks_payload, target_path)
            if res.get("ok"):
                self.original_path = str(target_path)
                self.working_path = None
                self.last_path = str(target_path)
                try:
                    self._cleanup_temp_files_for_original(Path(target_path))
                except Exception:
                    logger.exception("cleanup after save_as failed")
            return res
        except Exception:
            logger.exception("save_project_as failed")
            return {"error": "save_project_as failed"}

    def save_excel(self, payload):
        return self.save_project_as(payload)

    def quick_summary(self, payload):
        try:
            if pd is None:
                return {"error":"pandas not available"}
            df = pd.DataFrame(payload)
            if 'Start Date' not in df.columns: df['Start Date'] = pd.NaT
            if 'End Date' not in df.columns: df['End Date'] = pd.NaT
            if 'Status' not in df.columns: df['Status'] = ""
            if 'Task Name' not in df.columns: df['Task Name'] = ""
            for c in ['Start Date','End Date']:
                df[c] = pd.to_datetime(df[c], errors='coerce')
            df['Status'] = df['Status'].astype(str)
            today = pd.to_datetime(date.today())
            overdue = df[(df['End Date'].notna()) & (df['End Date'] < today) & (~df['Status'].str.lower().eq('done'))]
            due_today = df[(df['End Date'].notna()) & (df['End Date'] == today)]
            upcoming = df[(df['Start Date'].notna()) & (df['Start Date'] > today) & (df['Start Date'] <= today + pd.Timedelta(days=7))]
            def to_safe_list(dfall, cols):
                out=[]
                for _, r in dfall.iterrows():
                    item={}
                    for c in cols:
                        v = r.get(c,"")
                        if c in ('Start Date','End Date'): item[c] = date_to_iso(v)
                        else: item[c] = "" if pd.isna(v) else v
                    out.append(item)
                return out
            return {"summary":{"overdue":int(len(overdue)),"due_today":int(len(due_today)),"upcoming":int(len(upcoming)),"overdue_list":to_safe_list(overdue,['Task Name','End Date']),"due_today_list":to_safe_list(due_today,['Task Name']),"upcoming_list":to_safe_list(upcoming,['Task Name','Start Date'])}}
        except Exception:
            logger.exception("quick_summary failed")
            return {"error":"quick_summary failed"}

# ---------- logo embedding helpers ----------
def _recolor_svg_preserve_white(svg_text: str, target_color: str) -> str:
    try:
        def replace_attr(match):
            attr = match.group(1)
            value = match.group(2).strip()
            lv = value.lower()
            if lv in ('none', ''):
                return f'{attr}="{value}"'
            if re.match(r'^(#fff|#ffffff|white)$', lv) or 'rgb(255' in lv:
                return f'{attr}="{value}"'
            return f'{attr}="{target_color}"'
        svg_text = re.sub(r'(fill)\s*=\s*"([^"]*)"', replace_attr, svg_text, flags=re.IGNORECASE)
        svg_text = re.sub(r'(stroke)\s*=\s*"([^"]*)"', replace_attr, svg_text, flags=re.IGNORECASE)

        def replace_style(m):
            style = m.group(1)
            props = []
            for part in style.split(';'):
                if not part.strip():
                    continue
                if ':' not in part:
                    props.append(part)
                    continue
                k, v = part.split(':', 1)
                kk = k.strip().lower()
                vv = v.strip()
                lv = vv.lower()
                if kk in ('fill','stroke'):
                    if lv in ('none', '') or re.match(r'^(#fff|#ffffff|white)$', lv) or 'rgb(255' in lv:
                        props.append(f'{k}:{vv}')
                    else:
                        props.append(f'{k}:{target_color}')
                else:
                    props.append(f'{k}:{vv}')
            return 'style="' + ';'.join(props) + '"'
        svg_text = re.sub(r'style\s*=\s*"([^"]*)"', replace_style, svg_text, flags=re.IGNORECASE)
        return svg_text
    except Exception:
        return svg_text

def _write_small_png_from_bytes(png_bytes: bytes, out_path: Path, size=(72,72)):
    if Image is None:
        logger.debug("Pillow not available; cannot create small icon.")
        return False
    try:
        with Image.open(BytesIO(png_bytes)) as im:
            im = im.convert("RGBA")
            im.thumbnail(size, Image.LANCZOS)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            im.save(out_path, format="PNG")
            logger.info("Wrote small icon at %s", out_path)
            return True
    except Exception:
        logger.exception("Failed to create small icon from bytes")
        return False

def _write_small_png_from_file(src_path: Path, out_path: Path, size=(72,72)):
    if Image is None:
        logger.debug("Pillow not available; cannot create small icon.")
        return False
    try:
        with Image.open(str(src_path)) as im:
            im = im.convert("RGBA")
            im.thumbnail(size, Image.LANCZOS)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            im.save(out_path, format="PNG")
            logger.info("Wrote small icon at %s", out_path)
            return True
    except Exception:
        logger.exception("Failed to create small icon from file %s", src_path)
        return False

def prepare_logo_files(script_dir: Path) -> str:
    """
    Create/prepare local logo files and return the filename to use in HTML.
    The function now prefers writing the file next to the script and returning the filename
    (so injected HTML can reference it relative to index_injected.html).
    """
    svg_candidates = ["Logo.svg", "logo.svg", "LOGO.svg"]
    raster_candidates = ["Logo.png", "logo.png", "LOGO.png", "Logo.jpg", "logo.jpg", "LOGO.jpg"]
    target_svg = script_dir / "logo_injected.svg"
    target_raster = script_dir / "logo.png"
    small_icon = script_dir / "logo_small.png"
    logo_color = "#46C855"

    # Prefer SVG
    for c in svg_candidates:
        p = script_dir / c
        if p.exists():
            try:
                svg_text = p.read_text(encoding="utf-8")
                svg_text = svg_text.lstrip('\ufeff')
                recolored = _recolor_svg_preserve_white(svg_text, logo_color)
                try:
                    target_svg.write_text(recolored, encoding="utf-8")
                    logger.info("Wrote recolored SVG to %s", target_svg)
                except Exception:
                    logger.exception("Failed to write recolored SVG; falling back to source filename")
                    return p.name
                # Try to create small icon if possible
                if not small_icon.exists() and cairosvg and Image:
                    try:
                        png_bytes = cairosvg.svg2png(bytestring=recolored.encode('utf-8'))
                        _write_small_png_from_bytes(png_bytes, small_icon, size=(72,72))
                    except Exception:
                        logger.exception("Failed to rasterize SVG for small icon")
                return target_svg.name
            except Exception:
                logger.exception("Failed processing SVG %s", p)
                continue

    # Raster fallbacks (prefer read-bytes -> write, to avoid copy locks)
    for c in raster_candidates:
        p = script_dir / c
        if p.exists():
            try:
                # Try to read bytes from source (safer on locked files)
                read_ok = False
                data = None
                for attempt in range(3):
                    try:
                        data = p.read_bytes()
                        read_ok = True
                        break
                    except Exception as e:
                        logger.debug("Attempt %d: read_bytes failed for %s: %s", attempt+1, p, e)
                        time.sleep(0.12)
                if read_ok and data is not None:
                    try:
                        target_raster.write_bytes(data)
                        logger.info("Wrote raster logo by bytes to %s", target_raster)
                        if not small_icon.exists() and Image:
                            try:
                                _write_small_png_from_bytes(data, small_icon, size=(72,72))
                            except Exception:
                                logger.exception("Failed creating small icon from raster bytes")
                        return target_raster.name
                    except Exception as e:
                        logger.debug("write_bytes failed: %s", e)
                # Attempt shutil.copy2 (may fail on locked files)
                try:
                    shutil.copy2(str(p), str(target_raster))
                    logger.info("Copied raster logo to %s", target_raster)
                    if not small_icon.exists() and Image:
                        try:
                            _write_small_png_from_file(p, small_icon, size=(72,72))
                        except Exception:
                            logger.exception("Failed creating small icon from raster file")
                    return target_raster.name
                except Exception as e:
                    logger.debug("shutil.copy2 failed: %s", e)
                    # final fallback: return source filename
                    return p.name
            except Exception:
                logger.exception("Failed processing raster %s", p)
                continue

    # If running in PyInstaller onefile, attempt to extract from _MEIPASS
    if getattr(sys, 'frozen', False):
        meipass = Path(getattr(sys, '_MEIPASS', '.'))
        for c in svg_candidates:
            p = meipass / c
            if p.exists():
                try:
                    svg_text = p.read_text(encoding="utf-8")
                    svg_text = svg_text.lstrip('\ufeff')
                    recolored = _recolor_svg_preserve_white(svg_text, logo_color)
                    try:
                        target_svg.write_text(recolored, encoding="utf-8")
                        logger.info("Wrote recolored SVG from _MEIPASS to %s", target_svg)
                        return target_svg.name
                    except Exception:
                        logger.exception("Failed to write recolored SVG from _MEIPASS")
                        continue
                except Exception:
                    logger.exception("Failed processing SVG in _MEIPASS")
        for c in raster_candidates:
            p = meipass / c
            if p.exists():
                try:
                    raw = p.read_bytes()
                    try:
                        target_raster.write_bytes(raw)
                        logger.info("Wrote raster logo from _MEIPASS to %s", target_raster)
                        if not small_icon.exists() and Image:
                            try:
                                _write_small_png_from_bytes(raw, small_icon, size=(72,72))
                            except Exception:
                                logger.exception("Failed to create small icon from _MEIPASS raster")
                        return target_raster.name
                    except Exception:
                        logger.exception("Failed to write raster logo from _MEIPASS")
                except Exception:
                    logger.exception("Failed processing raster in _MEIPASS")
    return ""

# ---------- Windows monitor: detect minimize->restore and call JS redraw ----------
def _monitor_window_restore(window, window_title=None, poll_interval=0.25):
    """
    Background thread: polls OS window state and calls JS redraw when we detect a restore
    from minimized -> restored. Requires pywin32 (win32gui). This is best-effort: if win32gui
    is not available the function will be skipped.
    """
    if win32gui is None:
        logger.debug("win32gui not available; skipping window restore monitor")
        return

    title = window_title or getattr(window, "title", None) or "MBN Projects — Backend"
    prev_iconic = None
    logger.debug("Starting window restore monitor for title=%s", title)
    while True:
        try:
            hwnd = win32gui.FindWindow(None, title)
            if hwnd:
                is_iconic = bool(win32gui.IsIconic(hwnd))
                if prev_iconic is None:
                    prev_iconic = is_iconic
                # Detect transition from minimized -> restored
                if prev_iconic and not is_iconic:
                    logger.debug("Window restored detected for title=%s, triggering UI redraw", title)
                    try:
                        # call the exposed JS function (pywebview window object)
                        # The JS function is exposed as window.doFullRedrawWithRetries()
                        window.evaluate_js("window.doFullRedrawWithRetries && window.doFullRedrawWithRetries();")
                    except Exception:
                        try:
                            window.evaluate_js("doFullRedrawWithRetries && doFullRedrawWithRetries();")
                        except Exception:
                            logger.exception("Failed to call JS redraw from host monitor")
                prev_iconic = is_iconic
        except Exception:
            # swallow transient errors
            pass
        time.sleep(poll_interval)

# ---------- start webview ----------
def start():
    script_dir = Path(__file__).parent.resolve()
    index_path = script_dir / "index.html"
    if not index_path.exists():
        logger.error("index.html missing next to main.py in %s", script_dir)
        sys.exit(1)
    # Prepare logo and inject into index.html
    try:
        logo_src = prepare_logo_files(script_dir)
    except Exception:
        logger.exception("prepare_logo_files failed; continuing without logo")
        logo_src = ""
    html = index_path.read_text(encoding="utf-8")
    html = html.replace("{{LOGO_SRC}}", logo_src or "")
    api = Api()

    # write injected file for webview
    out = script_dir / "index_injected.html"
    try:
        out.write_text(html, encoding="utf-8")
    except Exception:
        import tempfile
        tmpdir = Path(tempfile.gettempdir())
        out = tmpdir / f"mbn_index_injected_{int(time.time())}.html"
        out.write_text(html, encoding="utf-8")

    def _remove_injected():
        try:
            if out.exists():
                out.unlink(missing_ok=True)
                logger.debug("Removed injected HTML: %s", out)
        except Exception:
            logger.exception("Failed removing injected HTML")
    try:
        atexit.register(_remove_injected)
    except Exception:
        logger.exception("Failed to register atexit cleanup for injected HTML")

    if webview is None:
        logger.info("pywebview not installed. Wrote %s. Open in a browser to test.", out)
        return
    try:
        url = out.as_uri()
        window = webview.create_window("MBN Projects — Backend", url=url, js_api=api, width=1200, height=820, resizable=True)

        # Start Windows monitor thread to detect minimize -> restore and call JS redraw
        try:
            if win32gui is not None:
                t = threading.Thread(target=_monitor_window_restore, args=(window, "MBN Projects — Backend"), daemon=True)
                t.start()
                logger.debug("Started window restore monitor thread")
        except Exception:
            logger.exception("Failed to start window restore monitor thread")

        webview.start(debug=False)
    except Exception:
        logger.exception("Failed to create webview window with file URL")
        return

if __name__ == "__main__":
    start()