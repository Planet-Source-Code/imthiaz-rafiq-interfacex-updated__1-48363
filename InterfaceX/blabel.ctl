VERSION 5.00
Begin VB.UserControl P_Blabel 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "blabel.ctx":0000
   Begin VB.Timer Controller 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   30
      Top             =   3150
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   825
      TabIndex        =   1
      Top             =   1590
      Width           =   2235
      Begin Interface.P_Bulb Bulb1 
         Height          =   180
         Index           =   0
         Left            =   30
         TabIndex        =   2
         Top             =   15
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         bulbflashcolor  =   16777215
      End
      Begin Interface.P_Bulb Bulb1 
         Height          =   180
         Index           =   1
         Left            =   255
         TabIndex        =   3
         Top             =   15
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         bulbflashcolor  =   16777215
      End
      Begin Interface.P_Bulb Bulb1 
         Height          =   180
         Index           =   2
         Left            =   465
         TabIndex        =   4
         Top             =   15
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         bulbflashcolor  =   16777215
      End
      Begin Interface.P_Bulb Bulb1 
         Height          =   180
         Index           =   3
         Left            =   690
         TabIndex        =   5
         Top             =   15
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         bulbflashcolor  =   16777215
      End
      Begin Interface.P_Bulb Bulb1 
         Height          =   180
         Index           =   4
         Left            =   900
         TabIndex        =   6
         Top             =   15
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         bulbflashcolor  =   16777215
      End
      Begin Interface.P_Bulb Bulb1 
         Height          =   180
         Index           =   5
         Left            =   1125
         TabIndex        =   7
         Top             =   15
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         bulbflashcolor  =   16777215
      End
      Begin Interface.P_Bulb Bulb1 
         Height          =   180
         Index           =   6
         Left            =   1335
         TabIndex        =   8
         Top             =   15
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         bulbflashcolor  =   16777215
      End
      Begin Interface.P_Bulb Bulb1 
         Height          =   180
         Index           =   7
         Left            =   1560
         TabIndex        =   9
         Top             =   15
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         bulbflashcolor  =   16777215
      End
      Begin Interface.P_Bulb Bulb1 
         Height          =   180
         Index           =   8
         Left            =   1770
         TabIndex        =   10
         Top             =   15
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         bulbflashcolor  =   16777215
      End
      Begin Interface.P_Bulb Bulb1 
         Height          =   180
         Index           =   9
         Left            =   1995
         TabIndex        =   11
         Top             =   15
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         bulbflashcolor  =   16777215
      End
   End
   Begin VB.Label Heading 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   1950
      TabIndex        =   0
      Top             =   525
      Width           =   630
   End
End
Attribute VB_Name = "P_Blabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private C_Bulb As Integer
Private C_Mode As Integer
Private C_Control As Boolean
Private Sub Timer1_Timer()

End Sub

Private Sub bulb1_onlightoff(Index As Integer)
If Index = Bulb1.UBound Then
    C_Bulb = 0
Else
    C_Bulb = Index + 1
End If
If C_Control <> False Then Controller.Enabled = True Else Controller.Enabled = False
End Sub

Private Sub Controller_Timer()
    Bulb1(C_Bulb).Onn = True
    Controller.Enabled = False
End Sub

Private Sub UserControl_Initialize()
    Heading.Caption = "Head"
    C_Bulb = 0
    C_Mode = 1
End Sub


Public Property Get Head() As String
    Head = Heading.Caption
End Property

Public Property Let Head(ByVal vNewValue As String)
    Heading.Caption = vNewValue
End Property

Public Property Get Border() As OLE_COLOR
    Border = Bulb1(0).BulbFlashColor
End Property

Public Property Let Border(ByVal vNewValue As OLE_COLOR)
    For i = Bulb1.LBound To Bulb1.UBound
        Bulb1(i).BulbFlashColor = vNewValue
    Next i
End Property


Private Sub UserControl_Resize()
    If UserControl.Width < Frame1.Width Then
        UserControl.Width = Frame1.Width
    End If
    Heading.Left = (UserControl.ScaleWidth / 2) - (Heading.Width / 2)
    Heading.Top = 0
    Frame1.Height = Bulb1(0).Height + 10
    Frame1.Width = Bulb1(Bulb1.UBound).Left + Bulb1(Bulb1.UBound).Width
    Frame1.Top = Heading.Height
    Frame1.Left = (UserControl.ScaleWidth / 2) - (Frame1.Width / 2)
    UserControl.Height = Frame1.Top + Frame1.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Head", Heading.Caption, "Head")
    Call PropBag.WriteProperty("Border", Bulb1(0).BulbFlashColor, vbWhite)
    Call PropBag.WriteProperty("Mode", C_Mode, 1)
    Call PropBag.WriteProperty("Control", C_Control, False)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Head = PropBag.ReadProperty("Head", "Head")
    Border = PropBag.ReadProperty("Border", vbWhite)
    Mode = PropBag.ReadProperty("Mode", 1)
    Control = PropBag.ReadProperty("Control", False)
End Sub


Public Property Get Mode() As Integer
    Mode = C_Mode
End Property

Public Property Let Mode(ByVal vNewValue As Integer)
    C_Mode = vNewValue
    Heading.ForeColor = getcolor(C_Mode, 255)
    For i = Bulb1.LBound To Bulb1.UBound
        Bulb1(i).Color = C_Mode
    Next i
End Property


Public Property Get Control() As Boolean
    Control = C_Control
End Property

Public Property Let Control(ByVal vNewValue As Boolean)
    C_Control = vNewValue
    If C_Control = True Then
        Controller.Enabled = True
    End If
End Property
