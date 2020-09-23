Attribute VB_Name = "Module2"
Global Exiter As Boolean
'----------------------Routines for Docking
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public dMouseStart As POINTAPI  ' Offset for window movement on screen
Public dWindowStart As POINTAPI ' Offset for window movement on screen
    '---------------------------------------------
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
    '---------------------------------------------


Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209

Public Const HWND_NOTOPMOST = -2

Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2

Public Const SWP_NOSIZE = &H1



Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long


Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_RESERVED = &HFF000000
Public Const SND_PURGE = &H40








Public Function LoadSound() As Boolean
    LoadSound = sndPlaySound(App.Path + "\Default.wav", SND_ASYNC Or SND_LOOP)
End Function
Public Function UnLoadSound() As Boolean
    UnLoadSound = sndPlaySound("", vbNull Or SND_PURGE)
End Function


Public Sub MeOnTop(Frm As Form)
    Dim lngWindowPosition As Long
    lngWindowPosition = SetWindowPos(Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Public Sub MeNotOnTop(Frm As Form)
    Dim lngWindowPosition As Long
    lngWindowPosition = SetWindowPos(Frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub


'----------docking routine

Public Sub DockingOnMouseMove(Index As Integer, Pixels As Long, Dockable As Boolean)

    Dim dMouseNow As POINTAPI
    Dim dWindowNow As POINTAPI
    GetCursorPos dMouseNow
    dWindowNow.X = Main.Holder(Index).Left \ Screen.TwipsPerPixelX
    dWindowNow.Y = Main.Holder(Index).Top \ Screen.TwipsPerPixelY
    If (dWindowNow.Y) < Pixels And Dockable Then
        If (dWindowStart.Y + (dMouseNow.Y - dMouseStart.Y)) < Pixels Then
            Main.Holder(Index).Top = 0
        Else
            Main.Holder(Index).Top = (dWindowStart.Y + (dMouseNow.Y - dMouseStart.Y)) * Screen.TwipsPerPixelY
        End If
        If (dWindowStart.X + (dMouseNow.X - dMouseStart.X)) < Pixels Then
            Main.Holder(Index).Left = 0
        Else
            Main.Holder(Index).Left = (dWindowStart.X + (dMouseNow.X - dMouseStart.X)) * Screen.TwipsPerPixelX
        End If
    ElseIf (dWindowNow.X) < Pixels And Dockable Then
        If (dWindowStart.X + (dMouseNow.X - dMouseStart.X)) < Pixels Then
            Main.Holder(Index).Left = 0
        Else
            Main.Holder(Index).Left = (dWindowStart.X + (dMouseNow.X - dMouseStart.X)) * Screen.TwipsPerPixelX
        End If
        If (dWindowStart.Y + (dMouseNow.Y - dMouseStart.Y)) < Pixels Then
            Main.Holder(Index).Top = 0
        Else
            Main.Holder(Index).Top = (dWindowStart.Y + (dMouseNow.Y - dMouseStart.Y)) * Screen.TwipsPerPixelY
        End If
    Else
        Main.Holder(Index).Left = (dWindowStart.X + (dMouseNow.X - dMouseStart.X)) * Screen.TwipsPerPixelX
        Main.Holder(Index).Top = (dWindowStart.Y + (dMouseNow.Y - dMouseStart.Y)) * Screen.TwipsPerPixelY
    End If
End Sub
Public Sub DockingOnMouseUp(Index As Integer, Pixels As Long, Dockable As Boolean)
    If Not (Dockable) Then Exit Sub
    If (Main.Holder(Index).Top / Screen.TwipsPerPixelY) < Pixels Then Main.Holder(Index).Top = 0
    If (Main.Holder(Index).Left / Screen.TwipsPerPixelX) < Pixels Then Main.Holder(Index).Left = 0
    If ((Main.Holder(Index).Top + Main.Holder(Index).Height) / Screen.TwipsPerPixelY) > CalDeskTopHeight - Pixels Then Main.Holder(Index).Top = (CalDeskTopHeight - (Main.Holder(Index).Height / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
    If ((Main.Holder(Index).Left + Main.Holder(Index).Width) / Screen.TwipsPerPixelX) > CalDeskTopWidth - Pixels Then Main.Holder(Index).Left = (CalDeskTopWidth - (Main.Holder(Index).Width / Screen.TwipsPerPixelX)) * Screen.TwipsPerPixelX
End Sub
Public Sub DockingOnMouseDown(Index As Integer, Dockable As Boolean)
    If Not (Dockable) Then Exit Sub
    GetCursorPos dMouseStart
    dWindowStart.X = Main.Holder(Index).Left \ Screen.TwipsPerPixelX
    dWindowStart.Y = Main.Holder(Index).Top \ Screen.TwipsPerPixelY
End Sub
Public Function CalDeskTopHeight()
    Dim hwnd As Long, rctemp As RECT
    Dim DeskTopHeight As Long
    Dim TaskBarHeight As Long

    hwnd = GetDesktopWindow()
    GetWindowRect hwnd, rctemp
    DeskTopHeight = rctemp.Bottom - rctemp.Top

    hwnd = FindWindow("Shell_TrayWnd", vbNullString)
    GetWindowRect hwnd, rctemp
    TaskBarHeight = rctemp.Bottom - rctemp.Top

    'If TaskBarHeight > DeskTopHeight Then
        CalDeskTopHeight = DeskTopHeight
   ' Else
    '    CalDeskTopHeight = DeskTopHeight - TaskBarHeight
   ' End If

End Function
Public Function CalDeskTopWidth()
    Dim hwnd As Long, rctemp As RECT
    Dim DesktopWidth As Long
    Dim TaskBarWidth As Long

    hwnd = GetDesktopWindow()
    GetWindowRect hwnd, rctemp
    DesktopWidth = rctemp.Right - rctemp.Left

    hwnd = FindWindow("Shell_TrayWnd", vbNullString)
    GetWindowRect hwnd, rctemp
    TaskBarWidth = rctemp.Right - rctemp.Left

    'If TaskBarWidth > DesktopWidth Then
        CalDeskTopWidth = DesktopWidth
    'Else
     '   CalDeskTopWidth = DesktopWidth - TaskBarWidth
   ' End If
End Function



