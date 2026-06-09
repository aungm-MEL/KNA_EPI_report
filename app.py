"""Run clean + long KNA pipelines from one Streamlit app."""

import os
import shutil
import subprocess
import sys
import tempfile
import io
import zipfile
from pathlib import Path

import streamlit as st


st.set_page_config(page_title="KNA EPI Pipeline", page_icon="💉", layout="wide")
st.title("💉 KNA EPI Pipeline")
st.caption(
    "Upload source child and Td workbooks, run build_kna_clean.py and build_kna_epi_long.py, "
    "then download both outputs."
)

base_dir = Path(__file__).resolve().parent
clean_script_candidates = [
    base_dir / "KNA_cleantoreport" / "build_kna_clean.py",
    base_dir / "build_kna_clean.py",
]
clean_script = next((p for p in clean_script_candidates if p.exists()), clean_script_candidates[0])
long_script = base_dir / "build_kna_epi_long.py"


def _script_looks_like_streamlit(script_path: Path) -> bool:
    try:
        text = script_path.read_text(encoding="utf-8", errors="ignore")
    except OSError:
        return False
    return "import streamlit as st" in text or "st.set_page_config(" in text

if not clean_script.exists() or not long_script.exists():
    clean_locations = "\n".join(f"- `{p}`" for p in clean_script_candidates)
    st.error(
        "Required scripts were not found. "
        "Make sure Streamlit Cloud **Main file path** is set to `KNA/app.py`.\n\n"
        f"- `app.py` is at: `{Path(__file__).resolve()}`\n"
        f"- Clean script search paths:\n{clean_locations}\n"
        f"- Selected clean script: `{clean_script}` — **{'FOUND' if clean_script.exists() else 'MISSING'}**\n"
        f"- Looking for long script:  `{long_script}` — **{'FOUND' if long_script.exists() else 'MISSING'}**"
    )
    st.stop()

if _script_looks_like_streamlit(clean_script):
    st.error(
        "Selected clean pipeline script appears to be a Streamlit UI file, not build logic.\n\n"
        f"- Selected path: `{clean_script}`\n"
        "- Expected: a data-processing script (pandas/openpyxl), not one containing `import streamlit as st`.\n"
        "Please verify your repository files and keep the clean script at `KNA_cleantoreport/build_kna_clean.py`."
    )
    st.stop()

with st.sidebar:
    st.header("Input Files")
    child_upload = st.file_uploader("KNA Child vaccination.xlsx", type=["xlsx", "xlsm"], key="child")
    td_upload = st.file_uploader("KNA Td Vaccination.xlsx", type=["xlsx", "xlsm"], key="td")

st.subheader("Upload Source Files")
col_u1, col_u2 = st.columns(2)
with col_u1:
    child_upload_main = st.file_uploader("Child source file", type=["xlsx", "xlsm"], key="child_main")
with col_u2:
    td_upload_main = st.file_uploader("Td source file", type=["xlsx", "xlsm"], key="td_main")

child_file = child_upload_main or child_upload
td_file = td_upload_main or td_upload

if child_file is None or td_file is None:
    st.info("Please upload both source files to continue.")
    st.stop()

st.success(f"Child: {child_file.name} | Td: {td_file.name}")

if not st.button("▶ Run Full Pipeline", type="primary", use_container_width=True):
    st.stop()

progress = st.progress(0, text="Starting pipeline...")
log_box = st.empty()
logs = []


def push_log(msg: str):
    logs.append(msg)
    log_box.code("\n".join(logs[-160:]), language="")


def run_step(cmd, cwd: Path, step_name: str, extra_env=None):
    push_log(f"Running: {step_name}")
    push_log(f"  cwd={cwd}")
    push_log(f"  cmd={' '.join(str(c) for c in cmd)}")
    env = dict(os.environ)
    if extra_env:
        env.update({k: str(v) for k, v in extra_env.items()})
        for k, v in extra_env.items():
            push_log(f"  env {k}={v}")
    res = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True, env=env)
    if res.stdout:
        for line in res.stdout.splitlines():
            push_log(f"  {line}")
    if res.stderr:
        for line in res.stderr.splitlines():
            push_log(f"  [stderr] {line}")
    if res.returncode != 0:
        stderr_tail = ""
        stdout_tail = ""
        if res.stderr:
            stderr_lines = [line for line in res.stderr.splitlines() if line.strip()]
            stderr_tail = " | ".join(stderr_lines[-3:])
        if res.stdout:
            stdout_lines = [line for line in res.stdout.splitlines() if line.strip()]
            stdout_tail = " | ".join(stdout_lines[-3:])

        detail = stderr_tail or stdout_tail or "No stderr/stdout captured"
        raise RuntimeError(
            f"{step_name} failed with exit code {res.returncode}. Details: {detail}"
        )


clean_bytes = None
long_bytes = None

try:
    with tempfile.TemporaryDirectory(prefix="kna_pipeline_") as tmp_root:
        tmp_root = Path(tmp_root)
        tmp_kna = tmp_root / "KNA"
        tmp_clean_dir = tmp_kna / "KNA_cleantoreport"
        tmp_clean_dir.mkdir(parents=True, exist_ok=True)

        progress.progress(8, text="Preparing temporary workspace...")

        shutil.copy2(long_script, tmp_kna / "build_kna_epi_long.py")
        shutil.copy2(clean_script, tmp_clean_dir / "build_kna_clean.py")

        tmp_clean_script = tmp_clean_dir / "build_kna_clean.py"
        tmp_long_script = tmp_kna / "build_kna_epi_long.py"

        (tmp_clean_dir / "KNA Child vaccination.xlsx").write_bytes(child_file.getvalue())
        (tmp_clean_dir / "KNA Td Vaccination.xlsx").write_bytes(td_file.getvalue())

        clean_out = tmp_clean_dir / "KNA_clean.xlsx"
        clean_env = {
            "KNA_CHILD_SRC": tmp_clean_dir / "KNA Child vaccination.xlsx",
            "KNA_TD_SRC": tmp_clean_dir / "KNA Td Vaccination.xlsx",
            "KNA_CLEAN_DST": clean_out,
        }

        progress.progress(30, text="Running build_kna_clean.py...")
        run_step([sys.executable, str(tmp_clean_script)], tmp_clean_dir, "build_kna_clean.py", clean_env)

        if not clean_out.exists():
            raise FileNotFoundError("KNA_clean.xlsx was not produced")

        # build_kna_epi_long.py prefers KNA_cleantoreport/KNA_clean.xlsx under its base dir.
        shutil.copy2(clean_out, tmp_kna / "KNA_clean.xlsx")

        long_out = tmp_kna / "KNA_EPI_long.xlsx"
        long_env = {
            "KNA_LONG_INPUT": clean_out,
            "KNA_LONG_OUTPUT": long_out,
        }

        progress.progress(60, text="Running build_kna_epi_long.py...")
        run_step([sys.executable, str(tmp_long_script)], tmp_kna, "build_kna_epi_long.py", long_env)

        if not long_out.exists():
            raise FileNotFoundError("KNA_EPI_long.xlsx was not produced")

        clean_bytes = clean_out.read_bytes()
        long_bytes = long_out.read_bytes()

    progress.progress(100, text="Done")
    push_log("Pipeline finished successfully.")

except Exception as exc:
    progress.empty()
    st.error(f"Pipeline failed: {exc}")
    st.stop()

st.success("Both outputs are ready.")

zip_buffer = io.BytesIO()
with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
    zip_file.writestr("KNA_clean.xlsx", clean_bytes)
    zip_file.writestr("KNA_EPI_long.xlsx", long_bytes)
zip_buffer.seek(0)

st.download_button(
    label="Download Both Files (.zip)",
    data=zip_buffer.getvalue(),
    file_name="KNA_pipeline_outputs.zip",
    mime="application/zip",
    use_container_width=True,
    type="primary",
)

dl1, dl2 = st.columns(2)
with dl1:
    st.download_button(
        label="Download KNA_clean.xlsx",
        data=clean_bytes,
        file_name="KNA_clean.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with dl2:
    st.download_button(
        label="Download KNA_EPI_long.xlsx",
        data=long_bytes,
        file_name="KNA_EPI_long.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
