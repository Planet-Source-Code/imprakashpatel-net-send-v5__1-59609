VERSION 5.00
Begin VB.Form frmNetSend 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Net Send Vr 5.4"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3060
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNetSend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      BackColor       =   &H00E0E0E0&
      Height          =   4410
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "Press Enter to send Message"
      Top             =   -45
      Width           =   3030
      Begin VB.TextBox txtMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   990
         Left            =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         ToolTipText     =   "Press Enter to send Message"
         Top             =   2715
         Width           =   2445
      End
      Begin VB.CheckBox chkClear 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Clear Message after sending"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   330
         TabIndex        =   11
         Top             =   3750
         Value           =   1  'Checked
         Width           =   2460
      End
      Begin VB.TextBox txtEnterName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   270
         TabIndex        =   5
         ToolTipText     =   "Press Enter to send Message"
         Top             =   465
         Width           =   2430
      End
      Begin VB.ListBox lstMachName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   1230
         Left            =   270
         Sorted          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Press Enter to send Message"
         Top             =   1200
         Width           =   2430
      End
      Begin VB.CommandButton cmdHide1 
         Cancel          =   -1  'True
         Caption         =   "&Hide"
         Height          =   315
         Left            =   1125
         TabIndex        =   7
         ToolTipText     =   "Press Esc to Hide "
         Top             =   3375
         Width           =   690
      End
      Begin VB.CommandButton cmdSend1 
         Caption         =   "&Send"
         Default         =   -1  'True
         Height          =   315
         Left            =   1665
         TabIndex        =   6
         ToolTipText     =   "Press Enter to send Message"
         Top             =   2970
         Width           =   690
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   285
         Left            =   480
         TabIndex        =   8
         ToolTipText     =   "Click to Exit Me"
         Top             =   3195
         Width           =   615
      End
      Begin VB.Image imgClearEvents 
         Height          =   240
         Left            =   2745
         Picture         =   "frmNetSend.frx":058A
         ToolTipText     =   "Clear Event Log Directly"
         Top             =   3045
         Width           =   240
      End
      Begin VB.Image imgShowEventViewer 
         Height          =   240
         Left            =   2745
         Picture         =   "frmNetSend.frx":0B14
         ToolTipText     =   "Open Event Viewer"
         Top             =   2730
         Width           =   240
      End
      Begin VB.Image cmdExit1 
         Height          =   255
         Left            =   195
         Picture         =   "frmNetSend.frx":0E9E
         ToolTipText     =   "Click to Exit Me"
         Top             =   4005
         Width           =   795
      End
      Begin VB.Image cmdSend 
         Height          =   255
         Left            =   2055
         Picture         =   "frmNetSend.frx":1241
         ToolTipText     =   "To Send Message (Press Enter)"
         Top             =   4005
         Width           =   795
      End
      Begin VB.Image cmdHide 
         Height          =   255
         Left            =   1110
         Picture         =   "frmNetSend.frx":15DE
         ToolTipText     =   "To Hide this Window (Press Escape)"
         Top             =   4005
         Width           =   795
      End
      Begin VB.Image lblDeleteMach 
         Height          =   240
         Left            =   2745
         Picture         =   "frmNetSend.frx":1982
         ToolTipText     =   "To Delete Machine Name"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image lblAddMach 
         Height          =   240
         Left            =   2730
         Picture         =   "frmNetSend.frx":1D0C
         ToolTipText     =   "To Add New Machine Name"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblOR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1215
         TabIndex        =   10
         Top             =   795
         Width           =   255
      End
      Begin VB.Label lblEnterName 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Machine &Name"
         Height          =   285
         Left            =   270
         TabIndex        =   4
         Top             =   225
         Width           =   2280
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Enter message here"
         Height          =   240
         Left            =   255
         TabIndex        =   2
         Top             =   2475
         Width           =   1965
      End
      Begin VB.Label lblMachName 
         BackStyle       =   0  'Transparent
         Caption         =   "Select &Machine name"
         Height          =   255
         Left            =   285
         TabIndex        =   0
         Top             =   960
         Width           =   2340
      End
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&systray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmNetSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function NetServerEnum Lib "netapi32.dll" (ByVal servername As String, _
      ByVal level As Long, buffer As Long, ByVal prefmaxlen As Long, entriesread As Long, _
      totalentries As Long, ByVal servertype As Long, ByVal domain As String, resumehandle As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32.dll" (BufPtr As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpszDest As String, ByVal lpszSrc As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_HIDE As Long = 0
Private Const ERROR_SUCCESS = 0
Private Const ERROR_MORE_DATA = 234
Private Const SV_TYPE_SERVER = &H2
Private Const SIZE_SI_101 = 24
Private Const NERR_Success As Long = 0&
Private Const NERR_BASE = 2100
Private Const NERR_NameNotFound = NERR_BASE + 173
Private Const NERR_NetworkError = NERR_BASE + 36
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_NOT_SUPPORTED = 50
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click

Private Type SERVER_INFO_101
        dwPlatformId As Long
        lpszServerName As Long
        dwVersionMajor As Long
        dwVersionMinor As Long
        dwType As Long
        lpszComment As Long
End Type
Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        uCallBackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private strSource As String
Private nid As NOTIFYICONDATA
Dim UniqueMCH As Collection
Const strMsg = "Net Send Vr 5.6"
Dim MaxMach As Integer
Dim OFName As OPENFILENAME
Private Sub Form_Load()
    
    Me.Caption = strMsg
    
    If Not ReadRegistrySettings Then
    
        LoadDefaultMachines
        
    End If
    
    ' The form must be fully visible before calling Shell_NotifyIcon
    Me.Show
    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Send Message" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid

    MaxMach = lstMachName.ListCount
End Sub
Private Function ReadRegistrySettings() As Boolean
    On Error GoTo errHandler
    Dim i As Integer
    Dim strMach As String
    'Check Registry
    If GetSetting("NetSend", "Developer", "Name") = "Prakash Patel" Then
    
        MaxMach = GetSetting("NetSend", "Developer", "MaxMach")
        If MaxMach > 0 Then
            ReadRegistrySettings = True
            For i = 0 To MaxMach - 1
                strMach = GetSetting("NetSend", "Machines", "Mach" & i)
                strMach = Trim$(strMach)
                
                If strMach <> vbNullString Then
                    lstMachName.AddItem strMach
                End If
                
            Next i
        Else
            ReadRegistrySettings = False
        End If
    Else
        ReadRegistrySettings = False
    End If
    
    Exit Function
errHandler:
    If Err.Number = 23 Then
        Err.Clear
        Resume Next
    Else
    MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & "frmNetSend.ReadRegistrySettings", _
        vbInformation, strMsg
    End If
End Function

Private Sub PriSubSendIt(LINK As String, Optional Args As String = "")
    
    Dim lRet As Long
    
    On Error Resume Next
    If LINK <> "" Then
        lRet = ShellExecute(0, "open", LINK, Args, App.Path, SW_HIDE)
        If lRet >= 0 And lRet <= 32 Then
            MsgBox "Error shelling to:" & LINK, 48, strMsg
        End If
    End If

End Sub

Private Sub cmdExit_Click()
    mPopExit_Click
End Sub

Private Sub cmdHide_Click()
    Me.WindowState = vbMinimized
End Sub


Private Sub cmdSend_Click()

    
    If UCase(Trim$(txtMessage.Text)) = "END" Then
        End
    End If
    
    If LenB(Trim$(txtEnterName.Text)) = 0 Then
        If lstMachName.Text <> vbNullString Then
            PriSubSendIt "net", "send " & Trim$(lstMachName.Text) & " " & Trim$(txtMessage.Text)
        Else
            MsgBox "Please Select / Enter Machine Name", vbInformation, strMsg
            'txtEnterName.Text = "Error !"
            Exit Sub
        End If
    Else
        If txtEnterName.Text <> vbNullString Then
            PriSubSendIt "net", "send " & Trim$(txtEnterName.Text) & " " & Trim$(txtMessage.Text)
        Else
            MsgBox "Please Select / Enter Machine Name", vbInformation, strMsg
            'txtEnterName.Text = "Error !"
            Exit Sub
        End If
    End If
    If chkClear.Value = vbChecked Then
        txtMessage.Text = vbNullString
    Else
    
    End If
    
    txtEnterName.Text = vbNullString
    
End Sub

Private Sub LoadDefaultMachines()

    Dim strServer As String
    Dim strDomain As String
    Dim lngLevel As Long
    Dim lngCounter As Long
    Dim lngBufPtr As Long
    Dim lngTempBufPtr As Long
    Dim lngPrefMaxLen As Long
    Dim lngEntriesRead As Long
    Dim lngTotalEntries As Long
    Dim lngServerType As Long
    Dim lngResumeHandle As Long
    Dim lngRet As Long
    Dim ServerInfo As SERVER_INFO_101
    Dim strMachineName As String

    lngLevel = 101
    lngBufPtr = 0
    lngPrefMaxLen = &HFFFFFFFF
    lngEntriesRead = 0
    lngTotalEntries = 0
    lngServerType = SV_TYPE_SERVER
    lngResumeHandle = 0
    
    Set UniqueMCH = New Collection
    
    Do
        lngRet = NetServerEnum(strServer, lngLevel, lngBufPtr, lngPrefMaxLen, lngEntriesRead, lngTotalEntries, lngServerType, strDomain, lngResumeHandle)
        If ((lngRet = ERROR_SUCCESS) Or (lngRet = ERROR_MORE_DATA)) And (lngEntriesRead > 0) Then
            lngTempBufPtr = lngBufPtr
            For lngCounter = 1 To lngEntriesRead
                RtlMoveMemory ServerInfo, lngTempBufPtr, SIZE_SI_101
                strMachineName = PriStrFunPointerToString(ServerInfo.lpszServerName)
                If UCase$(strMachineName) <> "AKASH" Then
                
                    CheckAndAdd strMachineName
                    
                End If
                
                lngTempBufPtr = lngTempBufPtr + SIZE_SI_101
            Next lngCounter
        Else
            MsgBox "NetServerEnum failed: " & lngRet, vbInformation, strMsg
        End If
        NetApiBufferFree (lngBufPtr)
    Loop While lngEntriesRead < lngTotalEntries
    

    AddotherMachines


End Sub
Sub AddotherMachines()
        CheckAndAdd "PRAKASH"
        CheckAndAdd "RAJ"
        CheckAndAdd "MANOJT"
        CheckAndAdd "SAMEERN"
        CheckAndAdd "NILESH"
        CheckAndAdd "DEEPAK"
        CheckAndAdd "ASHISH"
        CheckAndAdd "TEST5"
End Sub
Sub CheckAndAdd(ByVal strMach As String)
    On Error GoTo dubli
    'Check Unique and Add'
    UniqueMCH.Add strMach, strMach
    lstMachName.AddItem strMach
    
    Exit Sub
dubli:
    If Err.Number = 457 Then 'This key is already associated with an element of this collection
        Err.Clear
        Exit Sub
    End If
End Sub
Private Function PriStrFunPointerToString(ByVal vstrString As Long) As String

  Dim strFString As String
  Dim strSString As String
  Dim lngRet As Long
  
  strFString = String(1000, "*")
  lngRet = lstrcpyW(strFString, vstrString)
  strSString = (StrConv(strFString, vbFromUnicode))
  PriStrFunPointerToString = Left(strSString, InStr(strSString, Chr$(0)) - 1)
  
End Function



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      ' This procedure receives the callbacks from the System Tray icon.
      Dim lngResult As Long
      Dim lngMsg As Long
      
    ' The value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        lngMsg = X
    Else
        lngMsg = X / Screen.TwipsPerPixelX
    End If
    
    Select Case lngMsg
        Case WM_LBUTTONUP        '514 restore form window
            Me.WindowState = vbNormal
            lngResult = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            lngResult = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
            lngResult = SetForegroundWindow(Me.hwnd)
            'Me.mPopuMaxMachpMenu Me.mPopupSys
            PopupMenu Me.mPopupSys
    End Select
      
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        strSource = "Button"
        Me.WindowState = vbMinimized
    End If
    
End Sub

Private Sub Form_Resize()

    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
    
End Sub

Private Sub SaveRegistrySettings()
    On Error GoTo errHandler
    Dim i As Integer
    
    MaxMach = lstMachName.ListCount
    
    Call SaveSetting("NetSend", "Developer", "Name", "Prakash Patel")
    Call SaveSetting("NetSend", "Developer", "MaxMach", MaxMach)
    Call SaveSetting("NetSend", "Developer", "Version", strMsg)
    
    'Delete Old Data
    Call DeleteSetting("NetSend", "Machines")
    
    For i = 0 To MaxMach - 1
        Call SaveSetting("NetSend", "Machines", "Mach" & i, lstMachName.List(i))
    Next i

    Exit Sub
errHandler:
    If Err.Number = 5 Then
        Err.Clear
        Resume Next
    Else
    MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & "frmNetSend.SaveRegistrySettings", _
        vbInformation, strMsg
    End If
End Sub




Private Sub Image1_Click()

End Sub










Private Sub imgSendSMS_Click()
    frmVazu.Show
End Sub

Private Sub lblAddMach_Click()
    Dim strMach As String
    strMach = InputBox("Enter Machine Name", "Add New Machine Name")
    strMach = UCase(Trim$(strMach))
    If strMach <> vbNullString Then
        lstMachName.AddItem strMach
    End If
    MaxMach = lstMachName.ListCount '- 1
End Sub

Private Sub lblDeleteMach_Click()
    If lstMachName.ListCount > 0 Then
        If lstMachName.ListIndex <> -1 Then
            If MsgBox("Delete " & lstMachName.Text & " ? ", vbYesNo + vbInformation + vbDefaultButton2, strMsg) = vbYes Then
                lstMachName.RemoveItem lstMachName.ListIndex
            End If
        Else
            MsgBox "Please Select Machine to Delete", vbInformation, strMsg
        End If
    Else
        MsgBox "No Machine to Delete", vbInformation, strMsg
    End If
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Prakash Patel @ 2004-05", vbInformation, strMsg
End Sub

Private Sub mPopExit_Click()
    
    SaveRegistrySettings
    
    'called when user clicks the popup menu Exit command
    strSource = "Menu"
    Unload Me

End Sub

Private Sub mPopRestore_Click()

    Dim result
    'called when the user clicks the popup menu Restore command
    Me.WindowState = vbNormal
    result = SetForegroundWindow(Me.hwnd)
    Me.Show

End Sub
Private Sub Form_Unload(Cancel As Integer)

    'this removes the icon from the system tray
    If strSource = "Button" Then
        Cancel = True
        Exit Sub
    Else
        Shell_NotifyIcon NIM_DELETE, nid
    End If

End Sub

Private Sub cmdExit1_Click()
    cmdExit_Click
End Sub

Private Sub cmdHide1_Click()
    cmdHide_Click
End Sub

Private Sub cmdSend1_Click()
    cmdSend_Click
End Sub

Private Sub imgDownload_Click()
    On Error GoTo errHandler
    Dim obj As Scripting.FileSystemObject
    txtMessage.Text = Trim(txtMessage.Text)
    If txtMessage.Text <> vbNullString Then
        Set obj = New Scripting.FileSystemObject
        Dim strLocation As String
        If obj.FileExists(txtMessage.Text) Then
            'strLocation = InputBox("Save as ", "Where 2 Save ? ", "C:\")
            strLocation = LocateFile
            If strLocation <> vbNullString Then
                obj.CopyFile txtMessage.Text, strLocation
                MsgBox "File Saved to " & strLocation
            End If
        Else
            MsgBox "File Does not Exist", vbInformation, strMsg
        End If
    Else
        MsgBox "Please Enter the URL in Message Box " & vbCr & "Sample \\prakash\Shared\Test.exe", vbInformation, strMsg
        txtMessage.SetFocus
        'txtMessage.Text = "Sample \\prakash\Shared\Test.exe"
    End If
    Exit Sub
errHandler:
    MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & "frmNetSend.imgDownload_Click", _
        vbInformation, strMsg
End Sub
Private Function LocateFile() As String
    'Set the structure size
    OFName.lStructSize = Len(OFName)
    'Set the owner window
    OFName.hwndOwner = Me.hwnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Set the filet
    'OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'Create a buffer
    OFName.lpstrFile = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFile = 255
    'Create a buffer
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = "C:\"
    'Set the dialog title
    OFName.lpstrTitle = "Save File"
    'no extra flags
    OFName.flags = 0

    'Show the 'Save File'-dialog
    If GetSaveFileName(OFName) Then
        LocateFile = Trim$(OFName.lpstrFile)
    Else
        LocateFile = ""
    End If
End Function
Private Sub imgShowEventViewer_Click()
    Dim strParam As String
    
    If Trim(txtEnterName.Text) <> vbNullString Then
        strParam = Trim(txtEnterName.Text)
        txtEnterName.Text = ""
    ElseIf lstMachName.ListIndex <> -1 Then
        strParam = Trim(lstMachName.Text)
    Else
        strParam = ""
    End If
    
    'Shell "%SystemRoot%\system32\eventvwr.msc /s", vbNormalFocus
    ShellExecute Me.hwnd, "Open", "eventvwr.exe", strParam, "", 1
End Sub

Private Sub imgClearEvents_Click()
    If MsgBox("Clear Event Log ,Confirm ?", vbInformation + vbYesNo, strMsg) = vbYes Then
        If Trim(txtEnterName.Text) <> vbNullString Then
            txtEnterName.Text = Trim$(txtEnterName.Text)
        ElseIf lstMachName.ListIndex <> -1 Then
            ClearEventLogFile Trim(lstMachName.Text)
        Else
            ClearEventLogFile
        End If
    End If
End Sub
