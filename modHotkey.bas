Attribute VB_Name = "modHotkey"
'   Program: AutoDefrag 9.00
'Created by: Mark E. Twohey
'Created on: April 7th, 2006

'Begin declaration of Public API Constants
Public Const WM_CLOSE = &H10
Private Const SPI_SETSCREENSAVEACTIVE = 17

'Begin declaration of API Functions
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SwitchToThisWindow Lib "user32" (ByVal hWnd As Long, ByVal hWindowState As Long) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long

Sub Pause(duration)
    StartTime = Timer
    Do While Timer - StartTime < duration
        x = DoEvents()
    Loop
End Sub

Public Function ToggleScreenSaverActive(Active As Boolean) As Boolean
Dim lActiveFlag As Long
Dim retvaL As Long

lActiveFlag = IIf(Active, 1, 0)
retvaL = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, lActiveFlag, 0, 0)
ToggleScreenSaverActive = retvaL > 0
End Function


