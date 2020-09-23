VERSION 5.00
Begin VB.Form frmItemMaster 
   BackColor       =   &H00D1DECD&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2760
      Top             =   4440
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D1DECD&
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   7095
      Begin VB.TextBox txtOpeningStock 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3435
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Txtname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1950
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtItemCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   465
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   120
         Left            =   15
         Top             =   -30
         Width           =   7065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00D1DECD&
         Caption         =   "Opening Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4950
         TabIndex        =   14
         Top             =   465
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00D1DECD&
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3915
         TabIndex        =   13
         Top             =   465
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00D1DECD&
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   12
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00D1DECD&
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   735
         TabIndex        =   11
         Top             =   465
         Width           =   870
      End
   End
   Begin VB.CommandButton CmdMain 
      BackColor       =   &H00D1DECD&
      Caption         =   "&Main"
      Height          =   495
      Left            =   5760
      Picture         =   "frmItemMaster.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D1DECD&
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4680
      Picture         =   "frmItemMaster.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00D1DECD&
      Caption         =   "&Save"
      Height          =   495
      Left            =   3600
      Picture         =   "frmItemMaster.frx":028C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00D1DECD&
      Caption         =   "&New"
      Height          =   495
      Left            =   2640
      Picture         =   "frmItemMaster.frx":03FE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label LblDateTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "########"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4560
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "M S   F L E X G R I D  E X A M P L E   P R O G R A M  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   4200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I T E M  M A S T E R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   2160
      TabIndex        =   7
      Top             =   360
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmItemMaster.frx":0570
      Top             =   2520
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "frmItemMaster.frx":143A
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H009BAC8C&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   5
      Top             =   3360
      Width           =   7335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H009BAC8C&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmItemMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Clear frmItemMaster
TextLock frmItemMaster
CmdSave.Enabled = True
cmdNew.SetFocus
End Sub

Private Sub CmdMain_Click()
Me.Hide
frmMaster.Show
End Sub

Private Sub cmdNew_Click()
Clear frmItemMaster
TextUnlock frmItemMaster
txtItemCode.SetFocus
CmdSave.Enabled = False
End Sub

Private Sub CmdSave_Click()
Dim SLNo
SLNo = GetNewNo(" select max(slno)+1 from ItemMaster")
Update1 "ItemMaster", SLNo, txtItemCode, Txtname, txtRate, txtOpeningStock
MsgBox "New Item Sucessfully added ", vbInformation
cmdCancel_Click
End Sub

Private Sub Form_Load()
Clear frmItemMaster    ' Clear TextBoxes
TextLock frmItemMaster ' Lock Text boxes
End Sub


Private Sub Timer1_Timer()
LblDateTime.Caption = Time & " " & Format(Date, "DDDD")

End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
 If KeyAscii <> 13 Then
    If KeyAscii = 27 Then
        cmdNew.SetFocus
    Else
        KeyAscii = CheckCharecter(KeyAscii)
        KeyAscii = Asc(UCase(Chr(CheckCharecter(KeyAscii)))) '
    End If
 Else
   If Rs.State = 1 Then Rs.Close
   Rs.Open "select * from Itemmaster where Itmcode='" & txtItemCode & "'", myConection, adOpenKeyset, adLockOptimistic
       If Rs.EOF Then
            Txtname.SetFocus
       Else
            MsgBox " This ItemCode Already Exsist ! please choose another one ", vbCritical
            txtItemCode.Text = ""
            txtItemCode.SetFocus
       End If
 End If
 End Sub
Private Sub Txtname_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    If KeyAscii = 27 Then
       txtItemCode.SetFocus
    Else
       KeyAscii = CheckCharecter(KeyAscii)
       KeyAscii = Asc(UCase(Chr(CheckCharecter(KeyAscii))))
    End If
 Else
    txtRate.SetFocus
 End If
End Sub
Private Sub txtRate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    If KeyAscii = 27 Then
       Txtname.SetFocus
    Else
       KeyAscii = CheckNumeric(KeyAscii)
    End If
 Else
    txtOpeningStock.SetFocus
 End If
End Sub
Private Sub txtOpeningStock_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    If KeyAscii = 27 Then
       txtRate.SetFocus
    Else
       KeyAscii = CheckNumeric(KeyAscii)
    End If
 Else
   CmdSave.Enabled = True
   CmdSave.SetFocus
 End If
End Sub
