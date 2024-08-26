from typing import Optional

import psutil


def kill_process(pid: int) -> None:
    proc = psutil.Process(pid)
    proc.terminate()


def kill_all_processes(proc_name: str) -> None:
    for proc in psutil.process_iter():
        if proc_name in proc.name():
            try:
                proc.terminate()
            except (psutil.AccessDenied, psutil.NoSuchProcess):
                continue


def get_current_process_pid(proc_name: str) -> Optional[int]:
    return next(
        (p.pid for p in psutil.process_iter() if proc_name in p.name()),
        None,
    )
