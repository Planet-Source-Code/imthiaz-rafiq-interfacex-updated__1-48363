VERSION 5.00
Begin VB.UserControl P_Holder 
   BackColor       =   &H0080C0FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   6825
   ToolboxBitmap   =   "Holder.ctx":0000
   Begin VB.PictureBox Holder_Back 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   1755
      ScaleHeight     =   570
      ScaleWidth      =   615
      TabIndex        =   4
      Top             =   2475
      Width           =   615
      Begin VB.Shape Holder_Border 
         BackColor       =   &H00000000&
         BorderColor     =   &H00FF0000&
         Height          =   390
         Left            =   120
         Top             =   195
         Width           =   495
      End
   End
   Begin VB.Label Head_Space2 
      BackColor       =   &H000080FF&
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1635
      TabIndex        =   6
      Top             =   30
      Width           =   390
   End
   Begin VB.Label Head_Space1 
      BackColor       =   &H000080FF&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   5
      Top             =   15
      Width           =   195
   End
   Begin VB.Label Holder_Hide 
      BackColor       =   &H000080FF&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   3
      Top             =   30
      Width           =   210
   End
   Begin VB.Label Holder_Toggle 
      BackColor       =   &H000080FF&
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2055
      TabIndex        =   2
      Top             =   60
      Width           =   210
   End
   Begin VB.Label Holder_Head 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   " Holder  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   735
      TabIndex        =   1
      Top             =   45
      Width           =   855
   End
   Begin VB.Label Holder_Icon 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "ÿ"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   195
   End
End
Attribute VB_Name = "P_Holder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Holder_t_Status As Boolean
Dim Self_Resize As Boolean
Dim Top_ForeCOlor As OLE_COLOR
Dim Icons As Integer

Public Event HideClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)

Private Sub Head_Space1_DblClick()
    Holder_Toggle_Click
End Sub

Private Sub Head_Space1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Head_Space1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Head_Space1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Head_Space2_DblClick()
    Holder_Toggle_Click
End Sub

Private Sub Head_Space2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Head_Space2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Head_Space2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Holder_Head_Change()
    LoadTop
End Sub

Private Sub Holder_Head_DblClick()
    Holder_Toggle_Click
End Sub

Private Sub Holder_Head_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Holder_Head_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Holder_Head_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Holder_Hide_Click()
    RaiseEvent HideClick
End Sub


Private Sub Holder_Hide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Holder_Hide.ForeColor = vbWhite
End Sub

Private Sub Holder_Hide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Holder_Hide.ForeColor = Top_ForeCOlor
End Sub

Private Sub Holder_Icon_Click()
    Main.Show
End Sub

Private Sub Holder_Icon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Holder_Icon.ForeColor = vbWhite
End Sub

Private Sub Holder_Icon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Holder_Icon.ForeColor = Top_ForeCOlor
End Sub

Private Sub Holder_Toggle_Click()
    If Holder_t_Status = True Then
        Self_Resize = True
        Holder_Toggle.Caption = "q"
        UserControl.Height = Holder_Head.Height
        Self_Resize = True
        UserControl.Width = Holder_Hide.Width + Holder_Hide.Left
        Holder_t_Status = False
    Else
        Holder_Toggle.Caption = "y"
        Self_Resize = True
        UserControl.Height = Holder_Back.ScaleHeight + Holder_Head.Height
        Self_Resize = True
        UserControl.Width = Holder_Border.Width + Holder_Border.Left
        Holder_t_Status = True
    End If
    LoadTop
End Sub

Private Sub Holder_Toggle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Holder_Toggle.ForeColor = vbWhite
End Sub

Private Sub Holder_Toggle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Holder_Toggle.ForeColor = Top_ForeCOlor
End Sub



Private Sub UserControl_EnterFocus()
    Holder_Head.FontItalic = False
    Holder_Head.FontBold = True
    Holder_Head.ForeColor = vbWhite
    LoadTop
End Sub

Private Sub UserControl_ExitFocus()
    Holder_Head.FontItalic = True
    Holder_Head.FontBold = False
    Holder_Head.ForeColor = Top_ForeCOlor
    LoadTop
End Sub

Private Sub UserControl_Initialize()
    Holder_t_Status = True
    Self_Resize = False
    Icons = 1
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Holder_Toggle_Click
End If
If KeyCode = 8 Then
    RaiseEvent HideClick
End If
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_Resize()
If Self_Resize = False Then
    LoadTop
    LoadHolderBack
    BorderLoad
    Self_Resize = False
End If
End Sub
Sub LoadTop()
    Holder_Icon.Top = 0
    Holder_Icon.Left = 0
        
    Head_Space1.Top = 0
    Head_Space1.Left = Holder_Icon.Left + Holder_Icon.Width
    
    Holder_Head.Top = 0
    Holder_Head.Left = Head_Space1.Left + Head_Space1.Width
    
    Head_Space2.Top = 0
    Head_Space2.Left = Holder_Head.Left + Holder_Head.Width
    
    Holder_Toggle.Top = 0
    Holder_Toggle.Left = Head_Space2.Left + Head_Space2.Width
    
    Holder_Hide.Top = 0
    Holder_Hide.Left = Holder_Toggle.Left + Holder_Toggle.Width
    If (Holder_Hide.Width + Holder_Hide.Left) > UserControl.Width Then
        UserControl.Width = Holder_Hide.Width + Holder_Hide.Left
    End If
    If (Holder_Hide.Height + Holder_Hide.Top) > UserControl.Height Then
        UserControl.Height = Holder_Hide.Height + Holder_Hide.Height + 250
    End If
End Sub
Sub BorderLoad()
    Holder_Border.Left = 0
    Holder_Border.Top = 0
    Holder_Border.Width = Holder_Back.Width
    Holder_Border.Height = Holder_Back.Height
End Sub
Sub LoadHolderBack()
    Holder_Back.Left = 0
    Holder_Back.Top = Holder_Icon.Top + Holder_Icon.Height
    Holder_Back.Width = UserControl.Width
    Holder_Back.Height = UserControl.Height - Holder_Icon.Height
End Sub


Public Property Get HolderIcon() As String
    HolderIcon = Holder_Icon.Caption
End Property

Public Property Let HolderIcon(ByVal vNewValue As String)
    Holder_Icon.Caption = vNewValue
End Property

Public Property Get HolderHead() As String
    HolderHead = Holder_Head.Caption
End Property

Public Property Let HolderHead(ByVal vNewValue As String)
    Holder_Head.Caption = vNewValue
End Property

Public Property Get HeadBackColor() As OLE_COLOR
    HeadBackColor = Holder_Head.BackColor
End Property

Public Property Let HeadBackColor(ByVal vNewValue As OLE_COLOR)
    Holder_Icon.BackColor = vNewValue
    Holder_Head.BackColor = vNewValue
    Holder_Toggle.BackColor = vNewValue
    Holder_Hide.BackColor = vNewValue
    Head_Space1.BackColor = vNewValue
    Head_Space2.BackColor = vNewValue
End Property
Public Property Get HeadForeColor() As OLE_COLOR
    HeadForeColor = Holder_Head.ForeColor
End Property

Public Property Let HeadForeColor(ByVal vNewValue As OLE_COLOR)
    Holder_Icon.ForeColor = vNewValue
    Holder_Head.ForeColor = vNewValue
    Holder_Toggle.ForeColor = vNewValue
    Holder_Hide.ForeColor = vNewValue
    Top_ForeCOlor = vNewValue
End Property
Public Property Get HolderBackColor() As OLE_COLOR
    HolderBackColor = Holder_Back.BackColor
End Property

Public Property Let HolderBackColor(ByVal vNewValue As OLE_COLOR)
    Holder_Back.BackColor = vNewValue
    UserControl.BackColor = vNewValue
End Property
Public Property Get HolderBorderColor() As OLE_COLOR
    HolderBorderColor = Holder_Border.BorderColor
End Property

Public Property Let HolderBorderColor(ByVal vNewValue As OLE_COLOR)
    Holder_Border.BorderColor = vNewValue
End Property
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("HeadForeColor", Top_ForeCOlor, vbBlack)
    Call PropBag.WriteProperty("HeadBackColor", Holder_Icon.BackColor, &H80C0FF)
    Call PropBag.WriteProperty("HolderBackColor", Holder_Back.BackColor, &HC0FFFF)
    
    Call PropBag.WriteProperty("HolderBorderColor", Holder_Border.BorderColor, vbBlack)
    
    Call PropBag.WriteProperty("HolderHead", Holder_Head.Caption, vbBlack)
    Call PropBag.WriteProperty("HolderIcon", Holder_Icon.Caption, "ÿ")
    Call PropBag.WriteProperty("IconSet", Icons, 1)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    HeadForeColor = PropBag.ReadProperty("HeadForeColor", vbBlack)
    HeadBackColor = PropBag.ReadProperty("HeadBackColor", &H80C0FF)
    HolderBackColor = PropBag.ReadProperty("HolderBackColor", &HC0FFFF)
    
    HolderBorderColor = PropBag.ReadProperty("HolderBorderColor", vbBlack)
    
    HolderHead = PropBag.ReadProperty("HolderHead", "Holder")
    HolderIcon = PropBag.ReadProperty("HolderIcon", "ÿ")
    IconSet = PropBag.ReadProperty("IconSet", 1)
End Sub



Public Sub Toggle()
    Holder_Toggle_Click
End Sub

Public Property Get IconSet() As Variant
    IconSet = Icons
End Property

Public Property Let IconSet(ByVal vNewValue As Variant)
    Icons = vNewValue
    Select Case Icons
        Case 0
            Holder_Icon.Font = "Webdings"
        Case 1
            Holder_Icon.Font = "Wingdings"
        Case 2
            Holder_Icon.Font = "Wingdings 2"
        Case 3
            Holder_Icon.Font = "Wingdings 3"
        Case Else
            Holder_Icon.Font = "Wingdings"
            Icons = 1
    End Select
End Property
