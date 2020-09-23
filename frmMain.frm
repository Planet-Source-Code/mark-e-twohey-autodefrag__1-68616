VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoDefrag By Mark E. Twohey"
   ClientHeight    =   3045
   ClientLeft      =   12000
   ClientTop       =   735
   ClientWidth     =   3135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3045
   ScaleMode       =   0  'User
   ScaleWidth      =   3135
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   2640
      Top             =   1980
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillColor       =   &H00E0E0E0&
      ForeColor       =   &H00E0E0E0&
      HasDC           =   0   'False
      Height          =   2260
      Left            =   440
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   2692.248
      ScaleMode       =   0  'User
      ScaleWidth      =   2265
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   53
      Width           =   2260
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      FillColor       =   &H00E0E0E0&
      ForeColor       =   &H00E0E0E0&
      HasDC           =   0   'False
      Height          =   2260
      Left            =   440
      Picture         =   "frmMain.frx":60F4
      ScaleHeight     =   2692.247
      ScaleMode       =   0  'User
      ScaleWidth      =   2152.453
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   53
      Width           =   2260
   End
   Begin AutoDefrag.isButton isButton1 
      Height          =   375
      Left            =   780
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "frmMain.frx":BEDE
      Style           =   10
      Caption         =   "isButton"
      iNonThemeStyle  =   10
      ShowFocus       =   -1  'True
      HighlightColor  =   14737632
      FontColor       =   255
      FontHighlightColor=   14737632
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttBackColor     =   14737632
      ttForeColor     =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   14737632
      UseMaskColor    =   -1  'True
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2640
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   2400
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H000000FF&
      Height          =   2385
      Left            =   -120
      TabIndex        =   4
      Top             =   0
      Width           =   3240
   End
   Begin VB.Label Version 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   0
      TabIndex        =   7
      Top             =   2820
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Program: AutoDefrag 9.00
'Created by: Mark E. Twohey
'Created on: April 7th, 2006

Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpwindowname As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (VersionInfo As OSVERSIONINFOEX) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
    wServicePackMajor   As Integer
    wServicePackMinor   As Integer
    wSuiteMask          As Integer
    bProductType        As Byte
    bReserved           As Byte
End Type

Private Sub Form_Load()
        frmMain.Show
                                                                                                                                                                                                frmMain.Caption = Chr$(65) & Chr$(117) & Chr$(116) & Chr$(111) & Chr$(68) & Chr$(101) & Chr$(102) & Chr$(114) & Chr$(97) & Chr$(103) & Chr$(32) & Chr$(66) & Chr$(121) & Chr$(32) & Chr$(77) & Chr$(97) & Chr$(114) & Chr$(107) & Chr$(32) & Chr$(69) & Chr$(46) & Chr$(32) & Chr$(84) & Chr$(119) & Chr$(111) & Chr$(104) & Chr$(101) & Chr$(121)
    On Error Resume Next
        frmMain.MousePointer = 11
        frmMain.isButton1.Caption = "Loading . . ."
        frmMain.isButton1.Enabled = False
        Version.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
        SetWindowPos FindWindow(vbNullString, "AutoDefrag By Mark E. Twohey"), -1, 0, 0, 0, 0, &H1 Or &H2
        ToggleScreenSaverActive (False)
End Sub

Private Sub isButton1_Click()
        SendKeys "%{A}+{S}", True
        Wver = fWindowsVersion
        If Wver = "Microsoft Windows 2000 " Then
            Label2.Caption = "Microsoft Windows 2000 "
            GetClassNameFromTitle
        Else
            GetClassNameFromTitle
            Label2.Caption = ""
        End If
        
        Call SendMessage(Label1.Caption, WM_CLOSE, 0&, 0&)
        ToggleScreenSaverActive (True)
        Unload Me
        End
End Sub

Private Sub Timer1_Timer() 'Starts the Program
    Dim Buffer#
    Dim Defr As Long
    Dim mycommandline As String
    
    On Error Resume Next
    frmMain.Timer1.Enabled = False
    mycommandline = Trim(UCase(Command()))
    
    If mycommandline <> "" Then
        DType = GetDriveType(mycommandline)
            If DType <> 3 Then
                MsgBox ("Command Line " & mycommandline & " is not a Valid Hard Drive." & vbTab & "Syntax: AutoDefrag <drive letter>:" & vbTab & "Example: AutoDefrag d:" & vbTab), vbOKOnly + vbCritical, "AutoDefrag Error!"
                Unload frmMain
                End
            ElseIf DType = 3 Then
                    Buffer = Shell(Environ("WinDir") & "\system32\mmc.exe " & Environ("WinDir") & "\system32\dfrg.msc " & mycommandline, vbNormalFocus)  'Run the Windows Defrag executable file with focus.
            End If

    Else
        Buffer = Shell(Environ("WinDir") & "\system32\mmc.exe " & Environ("WinDir") & "\system32\dfrg.msc c:", vbNormalFocus) 'Run the Windows Defrag executable file with focus.
    End If
    
TryAgain:
        'Restores the Defrag Window
        Pause 1
        ShowWindow FindWindow(vbNullString, "Disk Defragmenter"), 9
        'Sets the Defrag Window Location and Stay on Top
        SetWindowPos FindWindow(vbNullString, "Disk Defragmenter"), -1, (32), (20), (635), (465), 0
        'Puts Focus on the Defrag Window
        Defr = FindWindow("MMCMainFrame", vbNullString)
        MMCMainFrame& = SwitchToThisWindow(Defr, vbNormalFocus)

        If MMCMainFrame& <> 1 Then
            GoTo TryAgain
        Else
            Pause 2
            SendKeys "%{A}+{D}", True
        End If
    
        'Starts the Third Timer
        Timer3.Enabled = True
End Sub

Public Sub GetClassNameFromTitle()
    Dim sInput As String
    Dim hWnd As Long
    Dim lpClassName As String
    Dim nMaxCount As Long
    Dim lResult As Long

    ' pad the return buffer for GetClassName
    nMaxCount = 256
    lpClassName = Space(nMaxCount)

    ' Note: must be an exact match
    If Label2.Caption = "Microsoft Windows 2000 " Then
        Wver = "Microsoft Windows XP "
        Label2.Caption = ""
    Else
        Wver = fWindowsVersion
    End If
    
    If Wver = "Microsoft Windows XP " Then
        sInput = "Disk Defragmenter" 'Window Title
    ElseIf Wver = "Microsoft Windows 2000 " Then
        sInput = "Defragmentation Complete" 'Window Title
    End If
        
    hWnd = FindWindow(vbNullString, sInput)

    ' Get the class name of the window, again, no validation
    lResult = GetClassName(hWnd, lpClassName, nMaxCount)
    Label2.Caption = Left$(lpClassName, lResult)
    Label1.Caption = hWnd
    Label3.Caption = hWnd

    If Label1.Caption = 0 And Label3.Caption = 0 Then
        Label4.Caption = "STOP"
    Else
        Label4.Caption = "Disk Defragmenter"
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Picture2.Visible = False
    Picture1.Visible = True
End Sub

Private Sub Timer2_Timer() 'Checks to see if the program is finished.
    Do While Label4.Caption = "Disk Defragmenter"
        If Label2.Caption = "#32770" Then
            Call SendMessage(Label1.Caption, WM_CLOSE, 0&, 0&)
            Pause 1
            Wver = fWindowsVersion
                If Wver = "Microsoft Windows 2000 " Then
                    Label2.Caption = "Microsoft Windows 2000 "
                    GetClassNameFromTitle
                Else
                    GetClassNameFromTitle
                    Label2.Caption = ""
                End If
            Label4.Caption = "END"
            Call SendMessage(Label3.Caption, WM_CLOSE, 0&, 0&)
            DoEvents

        Else
            DoEvents

        End If
    DoEvents
        
        If Label4.Caption = "END" Then
            frmMain.Timer1.Enabled = False
            frmMain.Timer2.Enabled = False
            ToggleScreenSaverActive (True)
            Unload Me
            End
        Else
            GetClassNameFromTitle
        End If
    Loop
            GetClassNameFromTitle
End Sub

Private Sub Timer3_Timer() 'Checks to see if the program has started.
    Dim DefChild As Long
    Dim DefParent As Long
    Dim Wver As String
                                                                                                                                                                                                frmMain.Caption = Chr$(65) & Chr$(117) & Chr$(116) & Chr$(111) & Chr$(68) & Chr$(101) & Chr$(102) & Chr$(114) & Chr$(97) & Chr$(103) & Chr$(32) & Chr$(66) & Chr$(121) & Chr$(32) & Chr$(77) & Chr$(97) & Chr$(114) & Chr$(107) & Chr$(32) & Chr$(69) & Chr$(46) & Chr$(32) & Chr$(84) & Chr$(119) & Chr$(111) & Chr$(104) & Chr$(101) & Chr$(121)
    Wver = fWindowsVersion

        If Wver = "Microsoft Windows XP " Then
            DefParent = FindWindow("MMCMainFrame", "Disk Defragmenter")
            DefChild = FindWindowEx(DefParent, 0&, "MDIClient", vbNullString)
            DefChild = FindWindowEx(DefChild, 0&, "MMCChildFrm", vbNullString)
            DefChild = FindWindowEx(DefChild, 0&, "MMCViewWindow", vbNullString)
            DefChild = FindWindowEx(DefChild, 0&, "MMCOCXViewWindow", vbNullString)
            DefChild = FindWindowEx(DefChild, 0&, "AtlAxWinEx", vbNullString)
            DefChild = FindWindowEx(DefChild, 0&, "ATL:6D3AAC48", vbNullString)
            DefChild = FindWindowEx(DefChild, 0&, "Button", "Defragment")

        ElseIf Wver = "Microsoft Windows 2000 " Then
            DefParent = FindWindow("MMCMainFrame", "Disk Defragmenter")
            DefChild = FindWindowEx(DefParent, 0&, "MDIClient", vbNullString)
            DefChild = FindWindowEx(DefChild, 0&, "MMCChildFrm", vbNullString)
            DefChild = FindWindowEx(DefChild, 0&, "AfxFrameOrView42u", vbNullString)
            DefChild = FindWindowEx(DefChild, 0&, "AfxFrameOrView42u", vbNullString)
            DefChild = FindWindowEx(DefChild, 0&, "ATL:726E6A38", vbNullString)
            DefChild = FindWindowEx(DefChild, 0&, "Button", "Defragment")
        End If

    If EnableWindow(DefChild, 1) = 0 Then   'Not Running
        EnableWindow DefChild, 1
        Defr = FindWindow("MMCMainFrame", vbNullString)
        MMCMainFrame& = SwitchToThisWindow(Defr, vbNormalFocus)
        SendKeys "%{A}+{D}", True
    Else                                    'Running
        EnableWindow DefChild, 0
        frmMain.isButton1.Caption = "Stop and Exit"
        frmMain.isButton1.Enabled = True
        frmMain.MousePointer = 0
        Timer3.Enabled = False
        Timer2.Enabled = True
    End If
End Sub

Public Function fWindowsVersion() As String 'Checks to see what Version of Windows you are using.
Dim VerInfo          As OSVERSIONINFOEX
Dim bOsVersionInfoEx As Long
   
    On Error Resume Next
    
    '
    ' Decide which UDT to use.
    ' Try OSVERSIONINFOEX first.
    '
    VerInfo.dwOSVersionInfoSize = Len(VerInfo)
    bOsVersionInfoEx = GetVersionEx(VerInfo)
    
    If bOsVersionInfoEx = 0 Then
        VerInfo.dwOSVersionInfoSize = OSVERSIONINFOSIZE
        Call GetVersionEx(VerInfo)
    End If
    
    Select Case VerInfo.dwPlatformId
        Case VER_PLATFORM_WIN32_NT
            If VerInfo.dwMajorVersion < 4 Then
                fWindowsVersion = "Microsoft Windows NT 3 "
            ElseIf VerInfo.dwMajorVersion = 4 Then
                fWindowsVersion = "Microsoft Windows NT 4 "
            ElseIf VerInfo.dwMajorVersion = 5 And VerInfo.dwMinorVersion = 0 Then
                fWindowsVersion = "Microsoft Windows 2000 "
            ElseIf VerInfo.dwMajorVersion = 5 And VerInfo.dwMinorVersion = 1 Then
                fWindowsVersion = "Microsoft Windows XP "
            End If
    
        Case VER_PLATFORM_WIN32_WINDOWS
            If VerInfo.dwMajorVersion = 4 And VerInfo.dwMinorVersion = 0 Then
                fWindowsVersion = "Microsoft Windows 95 "
            ElseIf VerInfo.dwMajorVersion = 4 And VerInfo.dwMinorVersion = 10 Then
                fWindowsVersion = "Microsoft Windows 98 "
            ElseIf VerInfo.dwMajorVersion = 4 And VerInfo.dwMinorVersion = 90 Then
                fWindowsVersion = "Microsoft Windows Me "
            End If
    
        Case VER_PLATFORM_WIN32s
            fWindowsVersion = "Microsoft Win32s "
    End Select

End Function
