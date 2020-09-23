VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   13305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17565
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13305
   ScaleWidth      =   17565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Interface.P_Holder Holder 
      Height          =   2535
      Index           =   6
      Left            =   12825
      TabIndex        =   31
      Top             =   8295
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   4471
      HeadForeColor   =   14737632
      HeadBackColor   =   4210752
      HolderBackColor =   16777215
      HolderBorderColor=   16777215
      HolderHead      =   "Options"
      HolderIcon      =   "%"
      IconSet         =   2
      Begin VB.PictureBox OptionHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2310
         Left            =   25
         ScaleHeight     =   2310
         ScaleWidth      =   3270
         TabIndex        =   32
         Top             =   255
         Width           =   3270
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2190
            Left            =   45
            MultiLine       =   -1  'True
            TabIndex        =   34
            Text            =   "frmMain.frx":27A2
            Top             =   45
            Width           =   3180
         End
      End
   End
   Begin Interface.P_Holder Holder 
      Height          =   3105
      Index           =   5
      Left            =   240
      TabIndex        =   23
      Top             =   480
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   5477
      HeadForeColor   =   14737632
      HeadBackColor   =   4210752
      HolderBackColor =   0
      HolderBorderColor=   14737632
      HolderHead      =   "Menu"
      HolderIcon      =   "¥"
      IconSet         =   0
      Begin Interface.P_Button Menu 
         Height          =   405
         Index           =   5
         Left            =   15
         TabIndex        =   30
         Top             =   2685
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   714
         IconFile        =   ""
         Color1          =   4210752
         Color2          =   8421504
         Color3          =   16777215
         Color4          =   12632256
         HColor1         =   16777215
         HColor2         =   0
         BCaption        =   "Exit"
      End
      Begin Interface.P_Button Menu 
         Height          =   405
         Index           =   4
         Left            =   15
         TabIndex        =   29
         Top             =   2280
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   714
         IconFile        =   ""
         Color1          =   4210752
         Color2          =   8421504
         Color3          =   16777215
         Color4          =   12632256
         HColor1         =   16777215
         HColor2         =   0
         BCaption        =   "About Us"
      End
      Begin Interface.P_Button Menu 
         Height          =   405
         Index           =   6
         Left            =   15
         TabIndex        =   28
         Top             =   1875
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   714
         IconFile        =   ""
         Color1          =   4210752
         Color2          =   8421504
         Color3          =   16777215
         Color4          =   12632256
         HColor1         =   16777215
         HColor2         =   0
         BCaption        =   "Options"
      End
      Begin Interface.P_Button Menu 
         Height          =   405
         Index           =   3
         Left            =   15
         TabIndex        =   27
         Top             =   1470
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   714
         IconFile        =   ""
         Color1          =   4210752
         Color2          =   8421504
         Color3          =   16777215
         Color4          =   12632256
         HColor1         =   16777215
         HColor2         =   0
         BCaption        =   "Console"
      End
      Begin Interface.P_Button Menu 
         Height          =   405
         Index           =   2
         Left            =   15
         TabIndex        =   26
         Top             =   1065
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   714
         IconFile        =   ""
         Color1          =   4210752
         Color2          =   8421504
         Color3          =   16777215
         Color4          =   12632256
         HColor1         =   16777215
         HColor2         =   0
         BCaption        =   "Admin Chat"
      End
      Begin Interface.P_Button Menu 
         Height          =   405
         Index           =   1
         Left            =   15
         TabIndex        =   25
         Top             =   660
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   714
         IconFile        =   ""
         Color1          =   4210752
         Color2          =   8421504
         Color3          =   16777215
         Color4          =   12632256
         HColor1         =   16777215
         HColor2         =   0
         BCaption        =   "Activity"
      End
      Begin Interface.P_Button Menu 
         Height          =   405
         Index           =   0
         Left            =   15
         TabIndex        =   24
         Top             =   255
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   714
         IconFile        =   ""
         Color1          =   4210752
         Color2          =   8421504
         Color3          =   16777215
         Color4          =   12632256
         HColor1         =   16777215
         HColor2         =   0
         BCaption        =   "Status"
      End
   End
   Begin Interface.P_Holder Holder 
      Height          =   2475
      Index           =   1
      Left            =   13635
      TabIndex        =   0
      Top             =   4725
      Width           =   2235
      _ExtentX        =   3625
      _ExtentY        =   423
      HeadForeColor   =   14737632
      HeadBackColor   =   4210752
      HolderBackColor =   0
      HolderBorderColor=   14737632
      HolderHead      =   "Activity"
      IconSet         =   0
      Begin VB.PictureBox ActivityHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2190
         Left            =   30
         ScaleHeight     =   2190
         ScaleWidth      =   2175
         TabIndex        =   4
         Top             =   270
         Width           =   2175
         Begin VB.Timer ActivityTimer 
            Interval        =   1000
            Left            =   405
            Top             =   735
         End
         Begin Interface.P_Blabel Ani 
            Height          =   435
            Index           =   3
            Left            =   0
            TabIndex        =   5
            Top             =   1305
            Width           =   2175
            _ExtentX        =   3942
            _ExtentY        =   767
            Head            =   "Chat Server"
            Mode            =   4
         End
         Begin Interface.P_Blabel Ani 
            Height          =   435
            Index           =   2
            Left            =   0
            TabIndex        =   6
            Top             =   870
            Width           =   2175
            _ExtentX        =   3942
            _ExtentY        =   767
            Head            =   "Cache Server"
            Mode            =   5
         End
         Begin Interface.P_Blabel Ani 
            Height          =   435
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   435
            Width           =   2175
            _ExtentX        =   3942
            _ExtentY        =   767
            Head            =   "Internet"
            Mode            =   6
         End
         Begin Interface.P_Blabel Ani 
            Height          =   435
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   2175
            _ExtentX        =   3942
            _ExtentY        =   767
            Head            =   "Proxy Server"
         End
         Begin Interface.P_Blabel Ani 
            Height          =   435
            Index           =   4
            Left            =   0
            TabIndex        =   22
            Top             =   1740
            Width           =   2175
            _ExtentX        =   3942
            _ExtentY        =   767
            Head            =   "Firewall"
            Mode            =   3
         End
      End
      Begin VB.Timer Offer 
         Interval        =   1000
         Left            =   1650
         Top             =   2475
      End
   End
   Begin Interface.P_Holder Holder 
      Height          =   3960
      Index           =   4
      Left            =   11655
      TabIndex        =   20
      Top             =   135
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6985
      HeadForeColor   =   14737632
      HeadBackColor   =   4210752
      HolderBackColor =   0
      HolderBorderColor=   14737632
      HolderHead      =   "About Us"
      HolderIcon      =   "#"
      IconSet         =   0
      Begin VB.PictureBox AboutHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3630
         Left            =   45
         ScaleHeight     =   3630
         ScaleWidth      =   3450
         TabIndex        =   21
         Top             =   285
         Width           =   3450
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   3495
            Left            =   60
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   33
            Text            =   "frmMain.frx":28E6
            Top             =   60
            Width           =   3315
         End
      End
   End
   Begin Interface.P_Holder Holder 
      Height          =   2655
      Index           =   3
      Left            =   6300
      TabIndex        =   16
      Top             =   8535
      Width           =   5760
      _ExtentX        =   3625
      _ExtentY        =   423
      HeadForeColor   =   14737632
      HeadBackColor   =   4210752
      HolderBackColor =   0
      HolderBorderColor=   14737632
      HolderHead      =   "Console"
      HolderIcon      =   ")"
      IconSet         =   0
      Begin VB.PictureBox CommandHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   15
         ScaleHeight     =   2100
         ScaleWidth      =   5730
         TabIndex        =   18
         Top             =   255
         Width           =   5730
         Begin VB.TextBox CommandText 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2085
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   0
            Width           =   5715
         End
      End
      Begin VB.TextBox CommandInput 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   15
         TabIndex        =   17
         Top             =   2340
         Width           =   5730
      End
   End
   Begin Interface.P_Holder Holder 
      Height          =   2655
      Index           =   2
      Left            =   195
      TabIndex        =   11
      Top             =   8535
      Width           =   5760
      _ExtentX        =   4128
      _ExtentY        =   423
      HeadForeColor   =   14737632
      HeadBackColor   =   4210752
      HolderBackColor =   0
      HolderBorderColor=   14737632
      HolderHead      =   "Admin Chat"
      HolderIcon      =   "("
      IconSet         =   0
      Begin VB.PictureBox ChatHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   15
         ScaleHeight     =   2100
         ScaleWidth      =   5730
         TabIndex        =   12
         Top             =   255
         Width           =   5730
         Begin VB.TextBox ChatText 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2085
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   0
            Width           =   5715
         End
      End
      Begin VB.TextBox ChatInput 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   15
         TabIndex        =   15
         Top             =   2340
         Width           =   5730
      End
   End
   Begin VB.PictureBox ToggleStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   330
      ScaleHeight     =   525
      ScaleWidth      =   4905
      TabIndex        =   9
      Top             =   5175
      Visible         =   0   'False
      Width           =   4905
      Begin VB.Timer ToggleStatusTimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   30
         Top             =   1635
      End
      Begin VB.Label ToggleStatusIcon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   450
         Left            =   330
         TabIndex        =   14
         Top             =   60
         Width           =   360
      End
      Begin VB.Label ToggleStatusHead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   765
         TabIndex        =   10
         Top             =   135
         Width           =   4230
      End
      Begin VB.Shape ToggleStatusBorder 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   510
         Left            =   15
         Top             =   15
         Width           =   4890
      End
   End
   Begin Interface.P_Holder Holder 
      Height          =   5220
      Index           =   0
      Left            =   3435
      TabIndex        =   1
      Top             =   135
      Width           =   6840
      _ExtentX        =   13070
      _ExtentY        =   12568
      HeadForeColor   =   14737632
      HeadBackColor   =   4210752
      HolderBackColor =   0
      HolderBorderColor=   14737632
      HolderHead      =   "Status"
      HolderIcon      =   "N"
      IconSet         =   0
      Begin VB.PictureBox StatusHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   4905
         Left            =   30
         ScaleHeight     =   4905
         ScaleWidth      =   6765
         TabIndex        =   2
         Top             =   270
         Width           =   6765
         Begin Interface.P_Display Status 
            Height          =   4800
            Left            =   45
            TabIndex        =   3
            Top             =   45
            Width           =   6660
            _ExtentX        =   11748
            _ExtentY        =   8467
         End
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Oricol(4) As Integer
Dim ActiveHolder As Integer
Dim ThruControl As Boolean

Private Sub ActivityTimer_Timer()
Dim S(2) As String
    X = Int(Rnd * 3)
    S(0) = "ü"
    S(1) = "ý"
    S(2) = "þ"
    Holder(1).HolderIcon = S(X)
End Sub

Private Sub ChatInput_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    ChatText.Text = ChatText.Text + "Admin : " + ChatInput.Text + vbCrLf
    ChatText.SelStart = Len(ChatText.Text) - 1
    ChatInput.Text = ""
End If
End Sub

Private Sub CommandInput_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CommandText.Text = CommandText.Text + "Command : " + CommandInput.Text + vbCrLf
    CommandText.Text = CommandText.Text + "Command " + CommandInput.Text + " executed" + vbCrLf
    CommandText.SelStart = Len(CommandText.Text) - 1
    CommandInput.Text = ""
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
'        Shell "start mailto:imthiazrafiq@hotmail.com", vbHide
        End
    ElseIf KeyCode = 9 Then
            ActiveHolder = ActiveHolder + 1
            If ActiveHolder = Holder.Count Then ActiveHolder = 0
            ToggleStatus.Top = (Me.Height / 2) - (ToggleStatus.Height / 2)
            ToggleStatus.Left = (Me.Width / 2) - (ToggleStatus.Width / 2)
            ToggleStatusHead.Caption = Holder(ActiveHolder).HolderHead
            ToggleStatusHead.Left = (ToggleStatus.Width / 2) - (ToggleStatusHead.Width / 2)
            Icons = Holder(ActiveHolder).IconSet
            Select Case Icons
                Case 0
                    ToggleStatusIcon.Font = "Webdings"
                Case 1
                    ToggleStatusIcon.Font = "Wingdings"
                Case 2
                    ToggleStatusIcon.Font = "Wingdings 2"
                Case 3
                    ToggleStatusIcon.Font = "Wingdings 3"
                Case Else
                    ToggleStatusIcon.Font = "Wingdings"
            End Select
            ToggleStatusIcon.Caption = Holder(ActiveHolder).HolderIcon
            ToggleStatusIcon.Left = ToggleStatusHead.Left - (ToggleStatusIcon.Width + 100)
            ToggleStatusTimer.Enabled = True
            ToggleStatus.ZOrder vbBringToFront
            ToggleStatus.Visible = True
    Else
        'MsgBox KeyCode
    End If
End Sub


Private Sub Form_Load()
For i = Ani.LBound To Ani.UBound
    Ani(i).Control = True
    Oricol(i) = Ani(i).Mode
Next i
For i = Menu.LBound To Menu.UBound
    Menu(i).IconFile = App.Path + "\icon.ico"
Next i
End Sub

Private Sub Holder_GotFocus(Index As Integer)
    Holder(Index).ZOrder vbBringToFront
    ActiveHolder = Index
    If Index = 2 Then
  '     ChatInput.SetFocus
    End If
End Sub

Private Sub Holder_HideClick(Index As Integer)
    Holder(Index).Visible = False
    If Index = 5 Then
        'Shell "start mailto:imthiazrafiq@hotmail.com", vbHide
        End
    End If
End Sub

Private Sub Holder_KeyPress(Index As Integer, KeyAscii As Integer)
If Holder(Index).Visible = True Then
    Select Case Index
        Case 2
            ChatInput.Text = ChatInput.Text + Chr(KeyAscii)
            ChatInput.SelStart = Len(ChatInput.Text)
            ChatInput.SetFocus
        Case 3
            CommandInput.Text = CommandInput.Text + Chr(KeyAscii)
            CommandInput.SelStart = Len(CommandInput.Text)
            CommandInput.SetFocus
    End Select
End If
End Sub

Private Sub Menu_Click(Index As Integer)
Select Case Index
    Case 5
        End
    Case Else
        ActiveHolder = Index
        If Holder(Index).Visible = True Then
            Holder(Index).Visible = False
        Else
            Holder(Index).Visible = True
            Holder(Index).SetFocus
            Holder(Index).ZOrder vbBringToFront
        End If
End Select
End Sub

Private Sub Offer_Timer()
    Dim X As Integer
    Offer.Interval = 1 + Rnd * 500
    X = Int(Rnd * Ani.Count)
    If Ani(X).Control = True Then
        Ani(X).Control = False
        Ani(X).Mode = 0
    '    Status.AddText Gettext(X), &HE0E0E0
    Else
        Ani(X).Control = True
        Ani(X).Mode = Oricol(X)
        Status.AddText GettextA(X), &HE0E0E0
    End If
End Sub

Private Sub Holder_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        DockingOnMouseDown Index, True
    End If
End Sub

Private Sub Holder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        DockingOnMouseMove Index, 20, True
    End If
End Sub
Private Sub Holder_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DockingOnMouseUp Index, 20, True
End Sub


Private Sub P_Holder1_HideClick()

End Sub

Private Sub ToggleStatusTimer_Timer()
ToggleStatusTimer.Enabled = False
ToggleStatus.Visible = False
If Holder(ActiveHolder).Visible = False Then
    Holder(ActiveHolder).Visible = True
    Holder(ActiveHolder).SetFocus
Else
    Holder(ActiveHolder).SetFocus
End If
End Sub
