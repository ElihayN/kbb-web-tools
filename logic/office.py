"""Helpers that shell out to LibreOffice (soffice) headless for format
conversion and formula recalculation — mirrors what the original Claude
skills did with `soffice --headless --convert-to ...`.
"""
import shutil
import subprocess
import os


class OfficeError(RuntimeError):
    pass


def _find_soffice():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    raise OfficeError(
        "LibreOffice (soffice) לא נמצא במערכת. יש להתקין LibreOffice כדי "
        "שהכלי יוכל להמיר קבצים ל-PDF/XLS ולחשב נוסחאות."
    )


def convert(src_path: str, out_dir: str, target_format: str, filter_name: str | None = None, timeout: int = 120) -> str:
    """Convert src_path to target_format (e.g. 'pdf', 'xlsx', 'csv') using
    LibreOffice headless. Returns the path to the produced file.

    filter_name lets us pick an explicit export filter, e.g. 'MS Excel 97'
    to force a true legacy .xls (matching the weekly-sms App2U requirement).
    """
    soffice = _find_soffice()
    os.makedirs(out_dir, exist_ok=True)
    target = target_format if not filter_name else f"{target_format}:{filter_name}"
    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--convert-to", target,
        "--outdir", out_dir,
        src_path,
    ]
    # LibreOffice headless can be flaky about its user profile directory when
    # several conversions run back-to-back; give each call its own profile.
    profile_dir = os.path.join(out_dir, ".lo_profile")
    env = os.environ.copy()
    cmd = cmd[:1] + [f"-env:UserInstallation=file://{profile_dir}"] + cmd[1:]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, env=env)
    combined_output = f"{result.stdout}\n{result.stderr}"
    # LibreOffice can exit 0 while still failing to write (e.g. converting a
    # file onto itself, same source/target path) — it only tells you via a
    # "Error:" line on stdout, so a bare returncode check isn't enough.
    if result.returncode != 0 or "Error:" in combined_output:
        raise OfficeError(
            f"המרת LibreOffice נכשלה עבור {src_path} -> {target_format}: "
            f"{combined_output.strip()}"
        )
    base = os.path.splitext(os.path.basename(src_path))[0]
    produced = os.path.join(out_dir, f"{base}.{target_format}")
    if not os.path.exists(produced):
        raise OfficeError(
            f"ההמרה רצה אך הקובץ הצפוי לא נוצר: {produced}\n{result.stdout}"
        )
    return produced


def recalc_xlsx(path: str, out_dir: str) -> str:
    """Round-trip an xlsx through LibreOffice so formula results are baked
    in as cached values (openpyxl only ever writes the formula string).

    Converting a file onto itself (same source/target path) makes
    LibreOffice fail silently (exit code 0, no output written), so this
    always renders into a throwaway subdirectory first and then moves the
    recalculated file back over the original path.
    """
    tmp_dir = os.path.join(out_dir, "_recalc_tmp")
    os.makedirs(tmp_dir, exist_ok=True)
    produced = convert(path, tmp_dir, "xlsx")
    final_path = path if path.endswith(".xlsx") else os.path.splitext(path)[0] + ".xlsx"
    shutil.move(produced, final_path)
    shutil.rmtree(tmp_dir, ignore_errors=True)
    return final_path


def to_pdf(path: str, out_dir: str) -> str:
    return convert(path, out_dir, "pdf")


def xlsx_to_legacy_xls(path: str, out_dir: str) -> str:
    return convert(path, out_dir, "xls", filter_name="MS Excel 97")


def xls_to_xlsx(path: str, out_dir: str) -> str:
    return convert(path, out_dir, "xlsx")
