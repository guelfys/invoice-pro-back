import os
import uuid
import subprocess
import threading
import sys
from dataclasses import dataclass
from typing import Dict, Optional, List


@dataclass
class Job:
    id: str
    status: str  # queued|running|done|error
    log_path: str
    exit_code: Optional[int] = None
    error: Optional[str] = None


_jobs: Dict[str, Job] = {}


def get_job(job_id: str) -> Job:
    return _jobs[job_id]


def _is_frozen_exe() -> bool:
    return bool(getattr(sys, "frozen", False))

def _runtime_root(project_root: str) -> str:
    if _is_frozen_exe():
        return os.path.dirname(sys.executable)
    return project_root

def _build_cmd(python_exe: str, full_script: str, args: List[str]) -> List[str]:
    if _is_frozen_exe():
        script_name = os.path.basename(full_script).lower()
        mode = "prod" if "produccion" in script_name or "producción" in script_name else "demo"
        return [sys.executable, "--run-generator", mode, *args]
    return [python_exe, "-u", full_script, *args]


def create_job(
    project_root: str,
    script_path: str,
    jobs_dir: str,
    python_exe: str,
    script_args: Optional[List[str]] = None,
    env_overrides: Optional[Dict[str, str]] = None,
    job_metadata: Optional[Dict[str, str]] = None,
) -> Job:
    os.makedirs(jobs_dir, exist_ok=True)

    job_id = uuid.uuid4().hex[:10]
    log_path = os.path.join(jobs_dir, f"{job_id}.log")

    job = Job(id=job_id, status="queued", log_path=log_path)
    _jobs[job_id] = job

    full_script = script_path if os.path.isabs(script_path) else os.path.join(project_root, script_path)

    def _run():
        job.status = "running"

        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"

        if env_overrides:
            env.update(env_overrides)

        args = script_args or []

        cwd = _runtime_root(project_root)

        cmd = _build_cmd(python_exe, full_script, args)

        try:
            with open(log_path, "w", encoding="utf-8", newline="\n") as f:
                f.write(f"[JOB {job_id}] Ejecutando:\n")
                if job_metadata:
                    for key, value in job_metadata.items():
                        f.write(f"  {key.upper()}: {value}\n")
                f.write(f"  EXE/PY: {python_exe}\n")
                f.write(f"  SCRIPT: {full_script}\n")
                f.write(f"  CWD:    {cwd}\n")
                f.write(f"  ARGS:   {args}\n")
                f.write(f"  CMD:    {cmd}\n")
                f.write("\n")
                f.flush()

                proc = subprocess.Popen(
                    cmd,
                    cwd=cwd,
                    stdout=f,
                    stderr=subprocess.STDOUT,
                    env=env,
                )
                rc = proc.wait()

                job.exit_code = rc
                job.status = "done" if rc == 0 else "error"

        except Exception as e:
            job.status = "error"
            job.error = str(e)

    threading.Thread(target=_run, daemon=True).start()
    return job


def tail_log(job_id: str, max_bytes: int = 80_000) -> str:
    job = _jobs[job_id]
    if not os.path.exists(job.log_path):
        return ""

    size = os.path.getsize(job.log_path)
    with open(job.log_path, "rb") as f:
        if size > max_bytes:
            f.seek(-max_bytes, os.SEEK_END)
        data = f.read()
    return data.decode("utf-8", errors="replace")
