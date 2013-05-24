'------------------------------------------------------------------------------ 
Const appName = "FTP Download Utility" 
'------------------------------------------------------------------------------

Const defaultHostname = "ftp.example.com" 
Const defaultPort = 21 
Const defaultUsername = "mshannon" 
Const defaultPassword = "welcome1" 
Const defaultRemoteDir = "/home/mshannon" 
Const defaultRemoteFile = "*.zip"

' set this var to the fully qualified path of a local directory to prevent 
' directory selection dialog from being displayed 
defaultLocalDir = "" 
' defaultLocalDir = "C:\Documents and Settings\Administrator\Desktop"

' if useDefaultsExclusively = True, the default values above will be leveraged 
' as-is, meaning no override options will be prompted for. 
Const useDefaultsExclusively = False 
' Const useDefaultsExclusively = True

' if skipConfirmation = True, the download will be attempted without requesting 
' confirmation to commence. 
Const skipConfirmation = False 
' Const skipConfirmation = True

'------------------------------------------------------------------------------

hostname = GetNonEmptyValue(useDefaultsExclusively, defaultHostname, _ 
  "Enter FTP server remote hostname:", "Hostname")

port = GetNonEmptyValue(useDefaultsExclusively, defaultPort, _ 
  "Enter FTP server remote port:", "Port")

username = GetNonEmptyValue(useDefaultsExclusively, defaultUsername, _ 
  "Enter username:", "Username")

password = GetNonEmptyValue(useDefaultsExclusively, defaultPassword, _ 
  "Enter password:", "Password")

If Len(defaultLocalDir) > 0 Then 
  localDir = defaultLocalDir 
Else 
  Set shell = CreateObject( "WScript.Shell" ) 
  defaultLocalDir = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop" 
  Set shell = Nothing 
  localDir = ChooseDirectory(defaultLocalDir, "Local Download Directory") 
  TestNotEmpty localDir, "Local Directory" 
End If

remoteDir = GetNonEmptyValue(useDefaultsExclusively, defaultRemoteDir, _ 
  "Remote download directory:", "Remote Directory")

remoteFile = GetNonEmptyValue(useDefaultsExclusively, defaultRemoteFile, _ 
  "Remote file to download (wildcard ok):", "Remote File")

Msg = "You have requested to download " & remoteFile & " from ftp://" & _ 
  username & "@" & hostname & ":" & port & remoteDir & _ 
  vbCRLF & "to: " & localDir & _ 
  vbCRLF & _ 
  vbCRLF & "Note - This may take quite some time!" & _ 
  vbCRLF & _ 
  vbCRLF & "Click OK to start download."

' VB appears to evaluate all the "OR" conditions when using if t1 OR t2 then ... 
' hence, it does not stop testing the conditions after the first condition 
' it detects is true. Thus the silly logic below... 
If skipConfirmation Then 
  Download hostname, port, username, password, localDir, remoteDir, remoteFile 
ElseIf vbOK = MsgBox(Msg, vbOKCancel, appName) Then 
  Download hostname, port, username, password, localDir, remoteDir, remoteFile 
End If

'------------------------------------------------------------------------------

Function GetNonEmptyValue(useDefaultExclusively, defaultValue, prompt, dialogTitle)

  If useDefaultExclusively Then 
    value = defaultValue 
  Else 
    value = InputBox(prompt, dialogTitle, defaultValue) 
  End If

  TestNotEmpty value, dialogTitle 
  GetNonEmptyValue = value 
End Function

'------------------------------------------------------------------------------

Sub TestNotEmpty(value, description) 
  If Len(value) = 0 Then 
    MsgBox "ERROR: No value provided for " & description, vbExclamation, appName 
    wscript.quit 
  End If 
End Sub

'------------------------------------------------------------------------------

Function ChooseDirectory(initialDir, prompt) 
  Set objShell  = CreateObject( "Shell.Application" )

  options = &H10& 'show text field to type folder path 
  'options = 0    'don't show text field to type folder path

  Set objFolder = objShell.BrowseForFolder(0, prompt, options, initialDir)

  If objFolder Is Nothing Then 
    ChooseDirectory = "" 
  Else 
    ChooseDirectory = objFolder.Self.Path 
  End If

  Set objFolder = Nothing 
  Set objShell = Nothing 
End Function

'------------------------------------------------------------------------------

Sub Download(hostname, port, username, password, localDir, remoteDir, remoteFile)

  Set shell = CreateObject("WScript.Shell") 
  Set fso = CreateObject("Scripting.FileSystemObject")

  tempDir = shell.ExpandEnvironmentStrings("%TEMP%") 
  ' temporary script file supplied to Windows FTP client 
  scriptFile = tempDir & "\" & fso.GetTempName 
  ' temporary file to store standard output from Windows FTP client 
  outputFile = tempDir & "\" & fso.GetTempName

  'input script 
  script = script & "lcd " & """" & localDir & """" & vbCRLF 
  script = script & "open " & hostname & " " & port & vbCRLF 
  script = script & "user " & username & vbCRLF 
  script = script & password & vbCRLF 
  script = script & "cd " & """" & remoteDir & """" & vbCRLF 
  script = script & "binary" & vbCRLF 
  script = script & "prompt n" & vbCRLF 
  script = script & "mget " & """" & remoteFile & """" & vbCRLF 
  script = script & "quit" & vbCRLF

  Set textFile = fso.CreateTextFile(scriptFile, True) 
  textFile.WriteLine(script) 
  textFile.Close 
  Set textFile = Nothing

  ' bWaitOnReturn set to TRUE - indicating script should wait for the program to 
  ' finish executing before continuing to the next statement 
  shell.Run "%comspec% /c FTP -n -s:" & scriptFile & " > " & outputFile, 0, TRUE 
  Wscript.Sleep 500 
  ' open standard output temp file read only, failing if not present 
  Set textFile = fso.OpenTextFile(outputFile, 1, 0, -2) 
  results = textFile.ReadAll 
  textFile.Close 
  Set textFile = Nothing 
  If InStr(results, "550") > 0 And InStr(results, "226") Then 
    fso.DeleteFile(scriptFile) 
    fso.DeleteFile(outputFile) 
    Msg ="WARNING: Could not change to destination directory on host!" & _ 
      vbCRLF & "File(s) however appear to have been downloaded from default " & _ 
      "FTP directory associated with user on host." 
    MsgBox Msg, vbExclamation, appName

  ElseIf InStr(results, "226") > 0 Then 
    MsgBox "File(s) Downloaded Successfully.", vbInformation, appName 
    fso.DeleteFile(scriptFile) 
    fso.DeleteFile(outputFile) 
  Else 
    If InStr(results, "530") > 0 Then 
      Msg ="ERROR: Invalid Username/Password" 
    ElseIf InStr(results, "550") > 0 Then 
      Msg ="ERROR: Could not open file on host" 
    ElseIf InStr(results, "553") > 0 Then 
      Msg ="ERROR: Could not create file on host" 
    ElseIf InStr(results, "Unknown host") > 0 Then 
      Msg ="ERROR: Unknown host" 
    ElseIf InStr(results, "File not found") > 0 Then 
      Msg ="ERROR: Local Directory Not Found" 
    Else 
      Msg ="An ERROR may have occurred." 
    End If

    Msg = Msg & _ 
      vbCRLF & "Script file leveraged: " & scriptFile & _ 
      vbCRLF & "FTP Output file: " & outputFile & _ 
      vbCRLF & _ 
      vbCRLF & "Ensure the above files are manually deleted, as they may " & _ 
      "contain sensitive information!" 
    ' Wscript.Echo Msg 
    MsgBox Msg, vbCritical, appName 
  End If 
  Set shell = Nothing 
  Set fso = Nothing

End Sub