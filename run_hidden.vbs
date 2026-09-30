' run_update.cmd를 콘솔 창 없이(숨김) 실행한다.
' Task Scheduler(lotto-lab-update)가 .cmd를 직접 실행하면 검은 창이 떠서 이 파일을 거친다.
' 종료까지 기다렸다가(True) 종료코드를 그대로 넘겨 작업 결과가 스케줄러에 남게 한다.
Dim rc, fso
Set fso = CreateObject("Scripting.FileSystemObject")
rc = CreateObject("WScript.Shell").Run( _
    "cmd /c """ & fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "run_update.cmd") & """", 0, True)
WScript.Quit rc
