Attribute VB_Name = "Module1"
'***************************************************************************
'* Copyright (c) 2004 by Prakah Patel
'*
'* This software is the proprietary information of Pd Systems.
'* Use is subject to license terms.
'*
'* @author  Prakash Patel
'* @version 1.0
'* @date    31 March 2004
'*
'***************************************************************************

Option Explicit

Public hEventLog As Long, LogString As String, Ret As Long ', ELR As EVENTLOGRECORD
Public bBytes(1 To 1024) As Byte
'Dim Computername As String
Public BackupFlag As eBackup
Public strFldPath As String
Public strBkupPath As String
Public Const StrLogFile As String = "c:\EventLog.txt"
Public Declare Function OpenEventLog Lib "advapi32.dll" Alias "OpenEventLogA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Public Declare Function BackupEventLog Lib "advapi32.dll" Alias "BackupEventLogA" (ByVal hEventLog As Long, ByVal lpBackupFileName As String) As Long
Public Declare Function ClearEventLog Lib "advapi32.dll" Alias "ClearEventLogA" (ByVal hEventLog As Long, ByVal lpBackupFileName As String) As Long
Public Const MAX_COMPUTERNAME_LENGTH As Long = 31

Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public strLogFlag As String
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim strCmdParam As String
Enum eBackup
    Clear = 0
    ClearBackUp = 1
    BackupSave = 2
End Enum
Public Const strMsg As String = "Net Send Admin"
'Enum eLogType
'    System = 1
'    Security = 2
'    Application = 3
'End Enum
Public Function ClearEventLogFile(Optional ByVal strComputername As String)

On Error GoTo errHandler
    
'Command$
'1=EventLOg Ttype
'2= 0-clear,1-clearbacku,2-backupsave
    If strComputername = vbNullString Then
        strComputername = GetLocalComputerName
    End If
    
    strCmdParam = "\\" & Trim$(strComputername) & "\"
    
    strLogFlag = "system"
    BackupFlag = Clear
                   
        
    If fncOPenEvent Then
        If BackupFlag = Clear Then
            fncClearEvent ("False")
        ElseIf BackupFlag = BackupSave Then
            subSetBkupPath
            fncGetBackup
        ElseIf BackupFlag = ClearBackUp Then
            subSetBkupPath
            fncClearEvent ("True")
        End If
    End If
    
Exit Function
errHandler:
    ClearEventLogFile = False
    Err.Clear
End Function
Private Function GetLocalComputerName()
On Error GoTo errHandler
    Dim Computername
    Dim dwLen As Long

    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    Computername = String(dwLen, "X")       'Create a buffer
    GetComputerName Computername, dwLen     'Get the computer name

    GetLocalComputerName = Left(Computername, dwLen)    'get only the actual data

    Exit Function
errHandler:
    'WriteLog "  " & Err.Description & " SubComputerName"
    'WriteLog String(100, " ")
    Err.Clear
End Function
Function fncOPenEvent() As Boolean
On Error GoTo errHandler
 
    hEventLog = OpenEventLog(strCmdParam, strLogFlag)
    If hEventLog <> 0 Then
        fncOPenEvent = True
    End If
    Exit Function
errHandler:
    fncOPenEvent = False
    'WriteLog Err.Description & Err.Source
    Err.Clear
End Function
Function fncClearEvent(ByVal ClearBackUp As Boolean) As Boolean
On Error GoTo errHandler
Dim Ret As Long
    If ClearBackUp Then
         'it will clear eventlog as well as take backup.
          Ret = ClearEventLog(hEventLog, strBkupPath)
       '  Ret = ClearEventLog(hEventLog, "C:\WINDOWS\system32\APRemoteInstall\EventLogBackup\System_Backup1.log")
    Else
          'it will clear eventlog only.
          Ret = ClearEventLog(hEventLog, vbNullString)
    End If
    If Ret <> 0 Then
        fncClearEvent = True
    End If

    Exit Function
errHandler:
    fncClearEvent = False
    'WriteLog Err.Description & Err.Source
    Err.Clear
End Function
Function fncGetBackup() As Boolean
On Error GoTo errHandler
    Dim Ret As Long
    'Write the event log to a file
     Ret = BackupEventLog(hEventLog, strBkupPath)
     If Ret <> 0 Then
        fncGetBackup = True
     End If
    Exit Function
errHandler:
    fncGetBackup = False
    'WriteLog Err.Description & Err.Source
    Err.Clear

End Function
Sub subSetBkupPath()
On Error GoTo errHandler
    Dim Fld As Folder
    Dim winDirPath As String, Ret As Long
    'Create a buffer
    winDirPath = Space(255)
    'Get the system directory
    Ret = GetSystemDirectory(winDirPath, 255)
    'Remove all unnecessary chr$(0)'s
    winDirPath = Left$(winDirPath, Ret)
    strFldPath = winDirPath & "\APRemoteInstall\EventLogBackup\"

    'If Not Fso.FolderExists(strFldPath) Then
    '    Set Fld = Fso.CreateFolder(strFldPath)
    'End If
    If Trim$(UCase$(strLogFlag)) = "SYSTEM" Then
       strBkupPath = strFldPath & "System_Backup.Log"
    ElseIf Trim$(UCase$(strLogFlag)) = "APPLICATION" Then
       strBkupPath = strFldPath & "Application_Backup.Log"
    ElseIf Trim$(UCase$(strLogFlag)) = "SECURITY" Then
        strBkupPath = strFldPath & "Security_Backup.Log"
    End If
    
    'If Fso.FileExists(strBkupPath) Then
    '       Call Fso.DeleteFile(strBkupPath, True)
    'End If
    
Exit Sub
errHandler:

    'WriteLog Err.Description & Err.Source
Err.Clear

End Sub

   

