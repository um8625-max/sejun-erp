' Sejun ERP quick launcher
' Double-click this file to start the server (minimized) and open the browser.
' To stop the server later, look for the minimized window in the taskbar and close it.

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
objShell.CurrentDirectory = strPath

' Install/check required packages (hidden window, waits until finished)
objShell.Run "cmd /c python -m pip install -r requirements.txt", 0, True

' Start the server, minimized to the taskbar (not fully hidden, so it can still be opened/closed)
objShell.Run "cmd /k python -m uvicorn main:app --host 0.0.0.0 --reload", 2, False

' Give the server a few seconds to start, then open the browser
WScript.Sleep 3000
objShell.Run "http://localhost:8000"
