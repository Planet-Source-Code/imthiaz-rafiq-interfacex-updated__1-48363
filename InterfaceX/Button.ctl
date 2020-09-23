VERSION 5.00
Begin VB.UserControl P_Button 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   345
   KeyPreview      =   -1  'True
   MousePointer    =   1  'Arrow
   ScaleHeight     =   10
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   23
   Begin VB.PictureBox Holder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   390
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   240
      Width           =   3000
      Begin VB.Label XCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Button"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   525
         TabIndex        =   1
         Top             =   285
         Width           =   810
      End
      Begin VB.Image Icon 
         Enabled         =   0   'False
         Height          =   240
         Left            =   255
         Picture         =   "Button.ctx":0000
         Stretch         =   -1  'True
         Top             =   225
         Width           =   240
      End
      Begin VB.Shape IconOver 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   405
         Left            =   165
         Top             =   195
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Shape IconBack 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   405
         Left            =   135
         Top             =   255
         Width           =   360
      End
   End
   Begin VB.Timer OnClick 
      Left            =   4395
      Top             =   1365
   End
End
Attribute VB_Name = "P_Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CanFocus As Boolean
Dim IsFocus As Boolean
Dim IsOver As Boolean
Dim IconFileName As String
Dim FColor1 As OLE_COLOR
Dim FColor2 As OLE_COLOR

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long




Private Sub OnClick_Timer()
    IconOver.Visible = Not (IconOver.Visible)
    OnClick.Interval = OnClick.Interval - 10
    If OnClick.Interval = 0 Then
        If CanFocus = True Then
            If IsFocus = True Then
                IconOver.Visible = True
                XCaption.ForeColor = FColor2
            Else
                IconOver.Visible = False
                XCaption.ForeColor = FColor1
            End If
        ElseIf IsOver = True Then
            IconOver.Visible = True
            XCaption.ForeColor = FColor2
        Else
            IconOver.Visible = False
            XCaption.ForeColor = FColor1
        End If
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_Click()
    OnClick.Interval = 100
End Sub

Private Sub UserControl_GotFocus()
    If CanFocus = True Then
        IconOver.Visible = True
        IsFocus = True
        XCaption.ForeColor = FColor2
    End If
End Sub

Private Sub UserControl_Initialize()
    CanFocus = False
    IsFocus = False
    IconFileName = "C:\WINDOWS\WINUPD.ICO"
    FColor1 = vbBlack
    FColor2 = vbWhite
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    IconOver.Visible = False
    IsFocus = False
    XCaption.ForeColor = FColor1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    Dim ret As Long
    Dim rct As RECT

    If X < 0 Or X > ScaleWidth Or Y < 0 Or Y > ScaleHeight Then
        fins = False
        ret = ReleaseCapture()
        If CanFocus = True Then
            If IsFocus = True Then
                IconOver.Visible = True
                XCaption.ForeColor = FColor2
            Else
                IconOver.Visible = False
                XCaption.ForeColor = FColor1
            End If
        Else
            IconOver.Visible = False
            XCaption.ForeColor = FColor1
        End If
        IsOver = False
    Else
        If fins = False Then
            fins = True
            ret = SetCapture(UserControl.hwnd)
            IconOver.Visible = True
            XCaption.ForeColor = FColor2
        End If
        IsOver = True
    End If
    
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    If UserControl.Height < 400 Then
        UserControl.Height = 400
        Exit Sub
    End If
    Holder.Top = 0
    Holder.Left = 0
    Holder.Height = UserControl.ScaleHeight
    Holder.Width = UserControl.ScaleWidth
    IconBack.Top = 0
    IconBack.Left = 0
    IconBack.Height = Holder.ScaleHeight + 1
    Icon.Top = (Holder.Height / 2) - (Icon.Height / 2)
    Icon.Left = (IconBack.Width / 2) - (Icon.Width / 2)
    IconOver.Top = 1
    IconOver.Left = 1
    IconOver.Width = Holder.Width - 2
    IconOver.Height = Holder.Height - 2
    XCaption.Left = IconBack.Width + 5
    XCaption.Top = (Holder.Height / 2) - (XCaption.Height / 2)
    If ((XCaption.Left + XCaption.Width) * Screen.TwipsPerPixelX) > UserControl.Width Then
        UserControl.Width = (XCaption.Left + XCaption.Width + 3) * Screen.TwipsPerPixelX
    End If
End Sub


Public Property Get IconFile() As String
    IconFile = IconFileName
End Property

Public Property Let IconFile(ByVal vNewValue As String)
    IconFileName = vNewValue
    Icon.Picture = LoadPicture(vNewValue)
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("IconFile", IconFileName, "C:\WINDOWS\WINUPD.ICO")
    Call PropBag.WriteProperty("Color1", Holder.BackColor, &H80000005)
    Call PropBag.WriteProperty("Color2", IconBack.BackColor, &H80000005)
    Call PropBag.WriteProperty("Color3", IconOver.BackColor, &HC0C0C0)
    Call PropBag.WriteProperty("Color4", IconOver.BorderColor, &H808080)
    Call PropBag.WriteProperty("HColor1", FColor1, vbBlack)
    Call PropBag.WriteProperty("HColor2", FColor2, vbWhite)
    Call PropBag.WriteProperty("GetFocus", CanFocus, False)
    Call PropBag.WriteProperty("BCaption", XCaption.Caption, "Button")
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    IconFile = PropBag.ReadProperty("IconFile", "C:\WINDOWS\WINUPD.ICO")
    Color1 = PropBag.ReadProperty("Color1", &H80000005)
    Color2 = PropBag.ReadProperty("Color2", &H80000005)
    Color3 = PropBag.ReadProperty("Color3", &HC0C0C0)
    Color4 = PropBag.ReadProperty("Color4", &H808080)
    HColor1 = PropBag.ReadProperty("HColor1", vbBlack)
    HColor2 = PropBag.ReadProperty("HColor2", vbWhite)
    GetFocus = PropBag.ReadProperty("Getfocus", False)
    BCaption = PropBag.ReadProperty("BCaption", "Button")
End Sub

Public Property Get Color1() As OLE_COLOR
    Color1 = Holder.BackColor
End Property

Public Property Let Color1(ByVal vNewValue As OLE_COLOR)
    Holder.BackColor = vNewValue
End Property

Public Property Get Color2() As OLE_COLOR
    Color2 = IconBack.BackColor
End Property

Public Property Let Color2(ByVal vNewValue As OLE_COLOR)
    IconBack.BackColor = vNewValue
End Property

Public Property Get Color3() As OLE_COLOR
    Color3 = IconOver.BackColor
End Property

Public Property Let Color3(ByVal vNewValue As OLE_COLOR)
    IconOver.BackColor = vNewValue
End Property

Public Property Get Color4() As OLE_COLOR
    Color4 = IconOver.BorderColor
End Property

Public Property Let Color4(ByVal vNewValue As OLE_COLOR)
    IconOver.BorderColor = vNewValue
End Property

Public Property Get HColor1() As OLE_COLOR
    HColor1 = FColor1
End Property

Public Property Let HColor1(ByVal vNewValue As OLE_COLOR)
    FColor1 = vNewValue
    XCaption.ForeColor = FColor1
End Property

Public Property Get HColor2() As OLE_COLOR
    HColor2 = FColor2
End Property

Public Property Let HColor2(ByVal vNewValue As OLE_COLOR)
    FColor2 = vNewValue
End Property



Public Property Get GetFocus() As Boolean
    GetFocus = CanFocus
End Property

Public Property Let GetFocus(ByVal vNewValue As Boolean)
    CanFocus = vNewValue
End Property

Public Property Get BCaption() As String
    BCaption = XCaption.Caption
End Property

Public Property Let BCaption(ByVal vNewValue As String)
    XCaption.Caption = vNewValue
    UserControl_Resize
End Property

