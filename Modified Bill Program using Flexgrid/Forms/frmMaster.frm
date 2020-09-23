VERSION 5.00
Begin VB.Form frmMaster 
   BackColor       =   &H00E1F2EE&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3795
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7695
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Imgexit 
      Height          =   510
      Left            =   1800
      MouseIcon       =   "frmMaster.frx":0000
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Image ImgSalesBill 
      Height          =   510
      Left            =   1800
      MouseIcon       =   "frmMaster.frx":030A
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Image ImgItemMaster 
      Height          =   495
      Left            =   1800
      MouseIcon       =   "frmMaster.frx":0614
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email : joshy_geo@hotmail.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   2235
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "website : www.joshygeo.tk"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "---------> All I want is to Give, because I have Received  <----------"
      Height          =   195
      Index           =   2
      Left            =   2880
      TabIndex        =   8
      Top             =   3240
      Width           =   4545
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "---------> So much help from others.  Thank You ! <----------"
      Height          =   195
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   3480
      Width           =   4545
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "---------> Use this as you like, no copyright, no restrictions <----------"
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   3000
      Width           =   4545
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   720
      Picture         =   "frmMaster.frx":091E
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   600
      Picture         =   "frmMaster.frx":20A0
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   2
      Left            =   3480
      TabIndex        =   5
      Top             =   2400
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Bill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Line Line3 
      BorderColor     =   &H009BAC8C&
      BorderWidth     =   25
      Index           =   0
      X1              =   1920
      X2              =   5760
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H009BAC8C&
      BorderWidth     =   25
      X1              =   1920
      X2              =   5760
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H009BAC8C&
      BorderWidth     =   25
      X1              =   1920
      X2              =   5760
      Y1              =   1560
      Y2              =   1560
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
      Height          =   915
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   7695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M S   F L E X G R I D  E X A M P L E   P R O G R A M   ( D E M O )"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   4815
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
      Height          =   1035
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Imgexit_Click()
Dim Msg, answer
   Msg = "Do you want to Exit?"
   answer = MsgBox(Msg, vbYesNo Or vbQuestion)
      If answer = vbYes Then
         End
      End If

End Sub

Private Sub ImgItemMaster_Click()
Me.Hide
frmItemMaster.Show
End Sub

Private Sub ImgSalesBill_Click()
Me.Hide
frmsalesbill.Show
End Sub
