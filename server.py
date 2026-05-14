"""
LL Sizing Report Flask backend.

Reads Calypso CMM measurement XLS files and provides APIs for the local web
dashboard. Filesystem access is intentionally limited to this project folder.
"""

import json
import os
import re
import subprocess

import xlrd
from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS


app = Flask(__name__, static_folder=".", static_url_path="")
CORS(app, resources={r"/api/*": {"origins": ["http://127.0.0.1:5050", "http://localhost:5050"]}})


SPEC_LIMITS = {
    "32mm": {"lower": 1.2610, "upper": 1.2635},
    "34mm": {"lower": 1.3396, "upper": 1.3421},
    "36mm": {"lower": 1.4180, "upper": 1.4205},
    "38mm": {"lower": 1.4965, "upper": 1.4990},
    "40mm": {"lower": 1.5760, "upper": 1.5785},
}
for size in SPEC_LIMITS:
    spec = SPEC_LIMITS[size]
    spec["mid"] = round((spec["lower"] + spec["upper"]) / 2, 7)


PATTERNS = {
    "L_Out1": re.compile(r"#\d+_ID[\d.]+_L\s*Out-1$", re.IGNORECASE),
    "L_Out2": re.compile(r"#\d+_ID[\d.]+_L\s*Out-2$", re.IGNORECASE),
    "L_In1": re.compile(r"#\d+_ID[\d.]+_L\s*In-1$", re.IGNORECASE),
    "L_In2": re.compile(r"#\d+_ID[\d.]+_L\s*In-2$", re.IGNORECASE),
    "R_Out1": re.compile(r"#\d+_ID[\d.]+_R\s*Out-1$", re.IGNORECASE),
    "R_Out2": re.compile(r"#\d+_ID[\d.]+_R\s*Out-2$", re.IGNORECASE),
    "R_In1": re.compile(r"#\d+_ID[\d.]+_R\s*In-1$", re.IGNORECASE),
    "R_In2": re.compile(r"#\d+_ID[\d.]+_R\s*In-2$", re.IGNORECASE),
}
REQUIRED_MEASUREMENTS = tuple(PATTERNS.keys())
REQUIRED_POSITIONS = ("AS_U", "AS_L", "DS_U", "DS_L")


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REGISTRY_FILE = os.path.join(BASE_DIR, "batch_registry.json")
REPORTS_DIR = os.path.join(BASE_DIR, "Reports")
SAFE_FILENAME_RE = re.compile(r"[^A-Za-z0-9_.-]+")

batch_registry = []


def is_within_base(path, base=BASE_DIR):
    """Return True when path resolves inside the configured base folder."""
    try:
        real_path = os.path.realpath(path)
        real_base = os.path.realpath(base)
        return os.path.commonpath([real_path, real_base]) == real_base
    except (OSError, ValueError):
        return False


def normalize_project_path(path):
    """Normalize and verify a user-supplied path."""
    norm_path = os.path.normpath(str(path or "").strip())
    if not norm_path:
        raise ValueError("Path is required")
    if not is_within_base(norm_path):
        raise ValueError("Path must be inside this project folder")
    return norm_path


def sanitize_report_filename(filename):
    """Return a predictable .html filename with unsafe characters replaced."""
    name = os.path.basename(str(filename or "Sizing_Report.html")).strip()
    name = SAFE_FILENAME_RE.sub("_", name)
    base, ext = os.path.splitext(name)
    if ext.lower() != ".html":
        ext = ".html"
    base = (base or "Sizing_Report")[:80]
    return f"{base}{ext}"


def build_report_path(filename):
    """Build a collision-free path that remains under Reports."""
    safe_filename = sanitize_report_filename(filename)
    os.makedirs(REPORTS_DIR, exist_ok=True)
    base, ext = os.path.splitext(safe_filename)
    filepath = os.path.join(REPORTS_DIR, safe_filename)
    counter = 1
    while os.path.exists(filepath):
        filepath = os.path.join(REPORTS_DIR, f"{base}_{counter}{ext}")
        counter += 1
    if not is_within_base(filepath, REPORTS_DIR):
        raise ValueError("Invalid report path")
    return filepath


def load_registry():
    """Load batch registry from JSON file."""
    global batch_registry
    if not os.path.isfile(REGISTRY_FILE):
        batch_registry = []
        return
    try:
        with open(REGISTRY_FILE, "r", encoding="utf-8") as f:
            batch_registry = json.load(f)
        print(f"  Loaded {len(batch_registry)} batches from registry")
    except (OSError, json.JSONDecodeError) as e:
        print(f"  Warning: Could not load registry: {e}")
        batch_registry = []


def save_registry():
    """Save batch registry to JSON file."""
    try:
        with open(REGISTRY_FILE, "w", encoding="utf-8") as f:
            json.dump(batch_registry, f, ensure_ascii=False, indent=2)
    except OSError as e:
        print(f"  Warning: Could not save registry: {e}")


def parse_xls_file(filepath):
    """Parse one Calypso XLS report and extract the required measurements."""
    wb = xlrd.open_workbook(filepath)
    sheet = wb.sheet_by_index(0)

    measurements = {}
    all_chars = {}
    warnings = []
    plan_name = ""
    part_no = ""
    date_val = ""

    try:
        plan_name = str(sheet.cell_value(4, 1)).strip()
    except (IndexError, ValueError, TypeError) as e:
        warnings.append(f"Could not read plan name: {e}")

    try:
        raw_pn = sheet.cell_value(7, 5)
        part_no = str(int(raw_pn)) if isinstance(raw_pn, float) else str(raw_pn)
    except (IndexError, ValueError, TypeError) as e:
        warnings.append(f"Could not read part number: {e}")

    try:
        date_val = sheet.cell_value(4, 3)
    except (IndexError, ValueError, TypeError):
        date_val = ""

    skipped_rows = 0
    for row in range(12, sheet.nrows):
        char_name = str(sheet.cell_value(row, 0)).strip()
        if not char_name or char_name == "Characteristic":
            continue
        try:
            all_chars[char_name] = {
                "actual": float(sheet.cell_value(row, 1)),
                "nominal": float(sheet.cell_value(row, 2)),
                "upper_tol": float(sheet.cell_value(row, 3)),
                "lower_tol": float(sheet.cell_value(row, 4)),
                "deviation": float(sheet.cell_value(row, 5)),
            }
        except (ValueError, IndexError, TypeError):
            skipped_rows += 1

    for char_name, char_data in all_chars.items():
        for key, pattern in PATTERNS.items():
            if pattern.match(char_name):
                measurements[key] = char_data["actual"]
                break

    missing_measurements = [key for key in REQUIRED_MEASUREMENTS if key not in measurements]
    if skipped_rows:
        warnings.append(f"Skipped {skipped_rows} non-numeric or malformed characteristic rows")
    if missing_measurements:
        warnings.append(f"Missing required measurements: {', '.join(missing_measurements)}")

    return {
        "plan": plan_name,
        "part_no": part_no,
        "date": date_val,
        "measurements": measurements,
        "all_characteristics": all_chars,
        "warnings": warnings,
    }


def compute_positions(measurements):
    """Compute the four position values from the eight measurements."""
    result = {}
    pairs = {
        "AS_U": ("L_Out1", "L_Out2"),
        "AS_L": ("L_In1", "L_In2"),
        "DS_U": ("R_Out1", "R_Out2"),
        "DS_L": ("R_In1", "R_In2"),
    }
    for position, (left_key, right_key) in pairs.items():
        left = measurements.get(left_key)
        right = measurements.get(right_key)
        if left is not None and right is not None:
            result[position] = round((left + right) / 2, 7)
    return result


def load_batch_data(folder_path):
    """Load all part data from a batch folder."""
    try:
        folder_path = normalize_project_path(folder_path)
    except ValueError as e:
        return None, [{"level": "error", "message": str(e)}]

    if not os.path.isdir(folder_path):
        return None, [{"level": "error", "message": "Folder not accessible"}]

    xls_files = sorted(f for f in os.listdir(folder_path) if f.lower().endswith(".xls"))
    parts = []
    warnings = []

    for index, filename in enumerate(xls_files, 1):
        filepath = os.path.join(folder_path, filename)
        try:
            parsed = parse_xls_file(filepath)
            positions = compute_positions(parsed["measurements"])
            file_warnings = list(parsed.get("warnings", []))
            missing_positions = [key for key in REQUIRED_POSITIONS if key not in positions]
            if missing_positions:
                file_warnings.append(f"Missing computed positions: {', '.join(missing_positions)}")

            parts.append({
                "index": index,
                "part_no": parsed["part_no"],
                "filename": filename,
                "measurements": parsed["measurements"],
                "positions": positions,
                "all_characteristics": parsed["all_characteristics"],
                "warnings": file_warnings,
            })
            for warning in file_warnings:
                warnings.append({"level": "warning", "file": filename, "message": warning})
        except Exception as e:
            print(f"Error parsing {filepath}: {e}")
            warnings.append({"level": "error", "file": filename, "message": str(e)})

    if len(parts) != len(xls_files):
        warnings.append({
            "level": "warning",
            "message": f"Parsed {len(parts)} of {len(xls_files)} .xls files",
        })

    return parts, warnings


def count_xls_files(folder_path):
    """Count .xls files in a directory inside the project folder."""
    try:
        folder_path = normalize_project_path(folder_path)
    except ValueError:
        return 0
    if not os.path.isdir(folder_path):
        return 0
    return len([f for f in os.listdir(folder_path) if f.lower().endswith(".xls")])


@app.route("/")
def index():
    response = send_from_directory(".", "index.html")
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.route("/style.css")
def css():
    return send_from_directory(".", "style.css")


@app.route("/api/batches")
def api_batches():
    return jsonify({"batches": batch_registry})


@app.route("/api/data/<path:batch_id>")
def api_data(batch_id):
    batch = next((b for b in batch_registry if b["id"] == batch_id), None)
    if batch is None:
        return jsonify({"error": "Batch not found"}), 404

    parts, warnings = load_batch_data(batch.get("path", ""))
    if parts is None:
        return jsonify({"error": "Folder not accessible", "warnings": warnings}), 404
    return jsonify({"batch_id": batch_id, "parts": parts, "warnings": warnings})


@app.route("/api/specs")
def api_specs():
    return jsonify({"specs": SPEC_LIMITS})


@app.route("/api/import", methods=["POST"])
def api_import():
    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON body"}), 400

    folder_path = str(data.get("path", "")).strip()
    batch_type = str(data.get("type", "")).strip()
    sizer_size = data.get("sizer_size")

    try:
        folder_path = normalize_project_path(folder_path)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    if not os.path.isdir(folder_path):
        return jsonify({"error": f"Folder not found: {folder_path}"}), 400
    if batch_type not in ("raw", "bushing_unsized", "bushing_sized"):
        return jsonify({"error": "Invalid type. Must be: raw, bushing_unsized, bushing_sized"}), 400

    if batch_type == "bushing_sized":
        try:
            sizer_size = float(sizer_size)
        except (TypeError, ValueError):
            return jsonify({"error": "Sizer size must be a number"}), 400
        if not 0.5 <= sizer_size <= 3.0:
            return jsonify({"error": "Sizer size is outside the expected inch range"}), 400
    else:
        sizer_size = None

    xls_count = count_xls_files(folder_path)
    if xls_count == 0:
        return jsonify({"error": "No .xls files found in folder"}), 400

    norm_path = os.path.normpath(folder_path)
    for batch in batch_registry:
        if os.path.normcase(os.path.normpath(batch["path"])) == os.path.normcase(norm_path):
            return jsonify({"error": "This folder is already imported"}), 409

    folder_name = os.path.basename(folder_path)
    if batch_type == "raw":
        label = "Raw Material"
    elif batch_type == "bushing_unsized":
        label = "Bushing ID (unsized)"
    else:
        label = f'Sized by {sizer_size:.5f}" sizer'

    date_match = re.match(r"(\d{6})", folder_name)
    batch_entry = {
        "id": folder_name,
        "folder": folder_name,
        "path": norm_path,
        "type": batch_type,
        "sizer_size": sizer_size,
        "label": label,
        "date_code": date_match.group(1) if date_match else "",
        "file_count": xls_count,
    }

    batch_registry.append(batch_entry)
    save_registry()
    print(f"  [IMPORT] {batch_type}: {folder_name} ({xls_count} files)")
    return jsonify({"ok": True, "batch": batch_entry}), 201


@app.route("/api/batch/<path:batch_id>", methods=["DELETE"])
def api_delete_batch(batch_id):
    global batch_registry
    before = len(batch_registry)
    batch_registry = [b for b in batch_registry if b["id"] != batch_id]
    if len(batch_registry) == before:
        return jsonify({"error": "Batch not found"}), 404
    save_registry()
    print(f"  [DELETE] {batch_id}")
    return jsonify({"ok": True})


@app.route("/api/browse")
def api_browse():
    target = request.args.get("path", "").strip()
    if not target:
        sizer_dir = os.path.join(BASE_DIR, "SIZER STUDY")
        target = sizer_dir if os.path.isdir(sizer_dir) else BASE_DIR

    try:
        target = normalize_project_path(target)
    except ValueError as e:
        return jsonify({"error": str(e), "path": target}), 400

    if not os.path.isdir(target):
        return jsonify({"error": "Path not found", "path": target}), 404

    entries = []
    try:
        for name in sorted(os.listdir(target)):
            full = os.path.join(target, name)
            if os.path.isdir(full) and is_within_base(full):
                entries.append({"name": name, "path": full, "xls_count": count_xls_files(full)})
    except PermissionError:
        return jsonify({"error": "Permission denied"}), 403

    parent_path = os.path.dirname(target)
    if not is_within_base(parent_path):
        parent_path = target
    return jsonify({"current_path": target, "parent_path": parent_path, "folders": entries})


@app.route("/api/save-report", methods=["POST"])
def api_save_report():
    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON body"}), 400

    html_content = data.get("html", "")
    if not isinstance(html_content, str) or not html_content.strip():
        return jsonify({"error": "HTML content is required"}), 400

    try:
        filepath = build_report_path(data.get("filename", "Sizing_Report.html"))
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(html_content)
    except (OSError, ValueError) as e:
        return jsonify({"error": str(e)}), 500

    actual_filename = os.path.basename(filepath)
    print(f"  [EXPORT] Saved report: {actual_filename}")
    return jsonify({
        "ok": True,
        "filename": actual_filename,
        "path": filepath,
        "folder": REPORTS_DIR,
    })


@app.route("/api/open-folder")
def api_open_folder():
    try:
        folder = normalize_project_path(request.args.get("path", ""))
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    if os.path.isdir(folder):
        subprocess.Popen(["explorer", folder])
        return jsonify({"ok": True})
    return jsonify({"error": "Invalid path"}), 400


load_registry()


if __name__ == "__main__":
    print("=" * 60)
    print("  LL Sizing Report Server")
    print("  http://127.0.0.1:5050")
    print("=" * 60)
    if batch_registry:
        for b in batch_registry:
            print(f"  [{b['type']:>15}] {b['folder']} ({b['file_count']} files)")
    else:
        print("  No batches imported yet. Use the UI to import folders.")
    print("=" * 60)
    app.run(host="127.0.0.1", port=5050, debug=False)
