## Excel Killer
## Script that force kills all active Excel Processes

import psutil
for proc in psutil.process_iter():
    if proc.name() == "EXCEL.EXE":
        proc.kill()
print("Excel closed using psutil")