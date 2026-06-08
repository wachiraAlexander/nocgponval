import os, subprocess, sys

APP_DIR = r"C:\Users\NOC\Downloads\GPON_PROJECT 1\GPON_PROJECT"

xml = f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2026-06-05T09:10:00</Date>
    <URI>\GPON App1</URI>
  </RegistrationInfo>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-21-2128426135-2682032236-4131866450-1001</UserId>
      <LogonType>Password</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
  </Settings>
  <Triggers>
    <LogonTrigger />
  </Triggers>
  <Actions Context="Author">
    <Exec>
      <Command>{os.path.join(APP_DIR, 'run_app.bat')}</Command>
      <WorkingDirectory>{APP_DIR}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>"""

path = os.path.join(os.environ["TEMP"], "gpon_task.xml")
with open(path, "w", encoding="utf-8") as f:
    f.write(xml)
print(f"Wrote: {path}")

r1 = subprocess.run(["schtasks", "/delete", "/TN", "GPON App1", "/F"], capture_output=True, text=True)
print("Delete:", r1.stdout.strip(), r1.stderr.strip())

r2 = subprocess.run(["schtasks", "/create", "/TN", "GPON App1", "/XML", path], capture_output=True, text=True)
print("Create:", r2.stdout.strip(), r2.stderr.strip())

r3 = subprocess.run(["schtasks", "/query", "/TN", "GPON App1", "/V", "/FO", "LIST"], capture_output=True, text=True)
for line in r3.stdout.splitlines():
    if "Task To Run" in line or "Start In" in line or "Status" in line:
        print(line)
