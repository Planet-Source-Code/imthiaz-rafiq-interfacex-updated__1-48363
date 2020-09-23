VERSION 5.00
Begin VB.UserControl P_Bulb 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "bulb.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2805
      Top             =   2655
   End
   Begin VB.Shape InnerBulb 
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1920
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape OuterBulb 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      Shape           =   5  'Rounded Square
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "P_Bulb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private BulbStatus As Boolean
Private BulbValue As Integer
Private BulbCOlor As Integer
Public Event onlightoff()
Private BulbOff As Integer
Private BulbSpeed As Integer
Private BulbOn As Integer

Private Sub Timer1_Timer()
If BulbStatus = True Then
    BulbValue = BulbValue + BulbSpeed
    If BulbValue > BulbOn Then
        BulbStatus = False
        OuterBulb.BorderStyle = 0
        BulbValue = BulbOff
        InnerBulb.FillColor = getcolor(BulbCOlor, BulbValue)
        InnerBulb.BorderColor = getcolor(BulbCOlor, BulbValue)
        Timer1.Enabled = False
        RaiseEvent onlightoff
    Else
        InnerBulb.FillColor = getcolor(BulbCOlor, BulbValue)
    End If
Else
    BulbValue = BulbValue - BulbSpeed
    If BulbValue = BulbOff Then
        BulbStatus = True
        OuterBulb.BorderStyle = 1
        BulbValue = BulbOff
        InnerBulb.FillColor = getcolor(BulbCOlor, BulbValue)
    End If
End If
End Sub

Private Sub UserControl_Initialize()
    BulbValue = 100
    OuterBulb.BorderStyle = 0
    BulbStatus = False
    BulbSpeed = 6
    BulbOff = 50
    BulbOn = 250 - BulbSpeed
End Sub


Private Sub UserControl_Resize()
    OuterBulb.Top = 0
    OuterBulb.Left = 0
    OuterBulb.Width = UserControl.ScaleWidth
    OuterBulb.Height = UserControl.ScaleHeight
    InnerBulb.Top = OuterBulb.Top + 1
    InnerBulb.Left = OuterBulb.Left + 1
    InnerBulb.Width = OuterBulb.Width - 2
    InnerBulb.Height = OuterBulb.Height - 2
End Sub

Public Property Get Onn() As Boolean
    Onn = BulbStatus
End Property

Public Property Let Onn(ByVal vNewValue As Boolean)
    BulbStatus = vNewValue
    BulbValue = BulbOff
    
    If vNewValue = True Then
        OuterBulb.BorderStyle = 1
        Timer1.Enabled = True
    End If
End Property

Public Property Get Color() As Integer
     Color = BulbCOlor
End Property

Public Property Let Color(ByVal vNewValue As Integer)
    BulbCOlor = vNewValue
     BulbValue = OffValue
     InnerBulb.FillColor = getcolor(BulbCOlor, BulbValue)
     InnerBulb.BorderColor = getcolor(BulbCOlor, BulbValue)
End Property

Public Property Get BulbFlashColor() As OLE_COLOR
    BulbFlashColor = OuterBulb.BorderColor
End Property

Public Property Let BulbFlashColor(ByVal vNewValue As OLE_COLOR)
    OuterBulb.BorderColor = vNewValue
End Property

Public Property Get Speed() As Integer
    Speed = BulbSpeed
End Property

Public Property Let Speed(ByVal vNewValue As Integer)
    BulbSpeed = vNewValue
End Property

Public Property Get OnValue() As Integer
   OnValue = BulbOn
End Property

Public Property Let OnValue(ByVal vNewValue As Integer)
    BulbOn = vNewValue
End Property


Public Property Get OffValue() As Integer
  OffValue = BulbOff
End Property

Public Property Let OffValue(ByVal vNewValue As Integer)
    BulbOff = vNewValue
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Speed", BulbSpeed, "10")
    Call PropBag.WriteProperty("Onvalue", BulbOn, "200")
    Call PropBag.WriteProperty("Offvalue", BulbOff, "50")
    Call PropBag.WriteProperty("bulbflashcolor", OuterBulb.BorderColor, "vbwhite")
    Call PropBag.WriteProperty("color", Color, "1")
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Speed = PropBag.ReadProperty("Speed", 10)
    OnValue = PropBag.ReadProperty("Onvalue", 200)
    OffValue = PropBag.ReadProperty("Offvalue", 50)
    BulbFlashColor = PropBag.ReadProperty("bulbflashcolor", vbWhite)
    Color = PropBag.ReadProperty("color", 1)
End Sub

