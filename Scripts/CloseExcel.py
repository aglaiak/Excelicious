## Close Excel
## Script that force kills all active Excel Processes

import psutil
def close_excel():
    for proc in psutil.process_iter():
        if proc.name() == "EXCEL.EXE":
            proc.kill()
