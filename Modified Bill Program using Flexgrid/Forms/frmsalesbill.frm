VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmsalesbill 
   BackColor       =   &H00E1F2EE&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5580
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   10755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1470
      Top             =   3540
   End
   Begin VB.TextBox txtInvoiceNo 
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
      ForeColor       =   &H80000010&
      Height          =   330
      Left            =   9180
      TabIndex        =   8
      Top             =   390
      Width           =   1515
   End
   Begin VB.TextBox txtTotal 
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
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   8700
      TabIndex        =   3
      Top             =   4770
      Width           =   1950
   End
   Begin VB.ComboBox cmbItmcode 
      Height          =   315
      Left            =   1905
      TabIndex        =   2
      Top             =   2505
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtEnter 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   150
      TabIndex        =   1
      Top             =   2490
      Visible         =   0   'False
      Width           =   1665
   End
   Begin MSFlexGridLib.MSFlexGrid MsfBill 
      Height          =   3360
      Left            =   -15
      TabIndex        =   0
      Top             =   1050
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   5927
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   13754061
      BackColorBkg    =   14807790
   End
   Begin VB.Image ImgCancel 
      Height          =   720
      Left            =   6645
      MouseIcon       =   "frmsalesbill.frx":0000
      MousePointer    =   99  'Custom
      Top             =   4695
      Width           =   780
   End
   Begin VB.Image ImgSave 
      Height          =   750
      Left            =   5985
      MouseIcon       =   "frmsalesbill.frx":030A
      MousePointer    =   99  'Custom
      Top             =   4680
      Width           =   645
   End
   Begin VB.Image ImgNew 
      Height          =   690
      Left            =   4740
      MouseIcon       =   "frmsalesbill.frx":0614
      MousePointer    =   99  'Custom
      Top             =   4710
      Width           =   1200
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
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
      Left            =   6090
      TabIndex        =   14
      Top             =   5190
      Width           =   450
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   6105
      Picture         =   "frmsalesbill.frx":091E
      Top             =   4695
      Width           =   480
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main"
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
      Left            =   6810
      TabIndex        =   13
      Top             =   5205
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6750
      Picture         =   "frmsalesbill.frx":15E8
      Top             =   4725
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
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
      Left            =   5310
      TabIndex        =   12
      Top             =   5160
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5310
      Picture         =   "frmsalesbill.frx":22B2
      Top             =   4650
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flex Grid Main "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1635
      TabIndex        =   11
      Top             =   4875
      Width           =   1200
   End
   Begin VB.Image ImgEditFlex 
      Height          =   510
      Left            =   555
      MouseIcon       =   "frmsalesbill.frx":2F7C
      MousePointer    =   99  'Custom
      Top             =   4710
      Width           =   3975
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
      Left            =   6210
      TabIndex        =   10
      Top             =   435
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8040
      TabIndex        =   9
      Top             =   435
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
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
      Left            =   7560
      TabIndex        =   7
      Top             =   4860
      Width           =   1020
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   165
      Picture         =   "frmsalesbill.frx":3286
      Top             =   285
      Width           =   480
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
      Left            =   810
      TabIndex        =   4
      Top             =   465
      Width           =   4200
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
      Left            =   -15
      TabIndex        =   5
      Top             =   0
      Width           =   10830
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00D7E1D0&
      BorderWidth     =   20
      X1              =   780
      X2              =   4125
      Y1              =   4965
      Y2              =   4980
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
      Height          =   1245
      Left            =   -15
      TabIndex        =   6
      Top             =   4335
      Width           =   10785
   End
End
Attribute VB_Name = "frmsalesbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Project Name:     Sales Bill                       |
' Module Name   :   BasMain                                |
' Purpose       :   Flexgrid Example program               |
' Author:           Joshy George                           |
' Start Date    :   31/08/2005 - 04:15 Pm                  |
'------------------------------------------------------------

Option Explicit
Dim gSlno, gItemCode, gItemname, gQty, gRate, gTotal, Inti
Dim Indx
Private Sub cmbItmcode_Change()
MsfBill.Text = cmbItmcode.Text
End Sub

Private Sub cmbItmcode_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler
If KeyAscii = 13 Then
   cmbItmcode.Visible = False
   MsfBill.TextMatrix(Indx, 1) = cmbItmcode.Text
   If Rs.State = 1 Then Rs.Close
   Rs.Open "select itmname,Rate from itemmaster where itmcode='" & cmbItmcode.Text & "'"
      If Not Rs.EOF Then
         MsfBill.TextMatrix(Indx, 2) = Rs!Itmname & ""
         MsfBill.TextMatrix(Indx, 3) = Rs!Rate & ""
         MsfBill.Col = 4
         ArrangeTextbox txtEnter
      Else
         MsgBox "Invalid Item code Please Check it ", vbCritical
         MsfBill.Col = 1
         ArrangeTextbox cmbItmcode
      End If
End If
Exit Sub
Err_Handler:
If Err.Number > 0 Then
  MsgBox Err.Description, vbCritical
End If
End Sub

Private Sub Form_Activate()
On Error GoTo Err_Handler
MsfBill.SetFocus
MsfBill.Row = 1
MsfBill.Col = 1
ArrangeTextbox cmbItmcode
Indx = 1
MsfBill.TextMatrix(Indx, 0) = Indx
txtInvoiceNo.Text = GetNewNo("select max(invoiceNo)+1 from sales")
Exit Sub
Err_Handler:
If Err.Number > 0 Then
  MsgBox Err.Description, vbCritical
End If
End Sub

Private Sub Form_Load()
MsfRefresh
FillCombo cmbItmcode, "select itmcode from ItemMaster"
End Sub
Private Sub MsfRefresh()
With MsfBill
      .Clear
      .Cols = 5
      .Rows = 2
      .FormatString = "^SL No | Item Code | Item Name | Rate  | Qty  | Total "
       gSlno = 0
       gItemCode = 1
       gItemname = 2
       gQty = 3
       gRate = 4
       gTotal = 5
       .Row = 0
       For Inti = 0 To .Cols - 1
          .Col = Inti
          .CellFontBold = True
       Next
       .ColWidth(gSlno) = 10 * 100
       .ColWidth(gItemCode) = 26 * 100
       .ColWidth(gItemname) = 26 * 100
       .ColWidth(gRate) = 15 * 100
       .ColWidth(gQty) = 15 * 100
       .ColWidth(gTotal) = 15 * 100
       .RowHeight(0) = 350
       .RowHeightMin = 350
End With
End Sub

Private Sub ArrangeTextbox(ctrl As Control)
  ctrl.Left = MsfBill.Left + MsfBill.CellLeft
  ctrl.Top = MsfBill.Top + MsfBill.CellTop
  ctrl.Text = MsfBill.Text
  ctrl.Width = MsfBill.ColWidth(MsfBill.Col) - 10
  If TypeOf ctrl Is TextBox Then
  ctrl.Height = MsfBill.RowHeight(MsfBill.Row) - 10
  End If
  ctrl.Visible = True
  ctrl.Text = ""
  ctrl.SetFocus
  ctrl.SelStart = 0
  ctrl.SelLength = Len(ctrl.Text)
End Sub

Private Sub ImgCancel_Click()
Me.Hide
frmMaster.Show
End Sub

Private Sub ImgNew_Click()
Clear frmsalesbill
txtInvoiceNo.Text = GetNewNo("select max(invoiceNo)+1 from sales")
MsfRefresh
MsfBill.SetFocus
MsfBill.Row = 1
MsfBill.Col = 1
ArrangeTextbox cmbItmcode
Indx = 1
MsfBill.TextMatrix(Indx, 0) = Indx
End Sub

Private Sub ImgSave_Click()
On Error GoTo Err_Handler
Dim I
Dim TrxType
TrxType = "S"
If MsgBox("Do you want to Save Bill", vbQuestion + vbYesNo + vbDefaultButton1, "Save Items") = vbYes Then
    For I = 1 To MsfBill.Row
     If Len(Trim(MsfBill.TextMatrix(I, 1))) = 0 Then
           MsgBox "Item Code. is Empty Please Enter"
           MsfBill.Row = I
           MsfBill.Col = 1
           Exit Sub
        End If
        If Len(Trim(MsfBill.TextMatrix(I, 4))) = 0 Then
           MsgBox "Qty. is Empty Please Enter"
           MsfBill.Row = I
           MsfBill.Col = 41
           Exit Sub
        End If
        If Len(Trim(MsfBill.TextMatrix(I, 3))) = 0 Then
           MsgBox "Rate is Empty Please Enter"
           MsfBill.Row = I
           MsfBill.Col = 3
           Exit Sub
        End If
        If Val(MsfBill.TextMatrix(I, 3)) = 0 Then
           MsgBox "Cheque Amount is Empty Please Enter"
           MsfBill.Row = I
           MsfBill.Col = 3
           Exit Sub
        End If
    Next
    For I = 1 To MsfBill.Row
        Update1 "Stock", MsfBill.TextMatrix(I, 1), MsfBill.TextMatrix(I, 4) * -1, TrxType, MsfBill.TextMatrix(I, 3)
        Update1 "Sales", txtInvoiceNo.Text, Format(Date, "dd-mmm-yyyy"), MsfBill.TextMatrix(I, 1), MsfBill.TextMatrix(I, 4), TrxType, MsfBill.TextMatrix(I, 5), txtTotal
    Next
    MsgBox "New Bill  details sucessfully Updated", vbInformation
    txtInvoiceNo.Text = GetNewNo("select max(invoiceNo)+1from sales")
    ImgNew_Click
End If
Exit Sub
Err_Handler:
If Err.Number > 0 Then
  MsgBox Err.Description, vbCritical
End If
End Sub

Private Sub MsfBill_Click()
  If MsfBill.Col = 1 Then
     MsfBill.Col = 1
     ArrangeTextbox cmbItmcode
  ElseIf MsfBill.Col = 2 Then
     MsfBill.Col = 2
     ArrangeTextbox txtEnter
  ElseIf MsfBill.Col = 3 Then
     MsfBill.Col = 3
     ArrangeTextbox txtEnter
  ElseIf MsfBill.Col = 4 Then
     MsfBill.Col = 4
     ArrangeTextbox txtEnter
  ElseIf MsfBill.Col = 5 Then
     MsfBill.Col = 5
     ArrangeTextbox txtEnter
  End If
End Sub

Private Sub Timer1_Timer()
LblDateTime.Caption = Time & " " & Format(Date, "DDDD")
End Sub

Private Sub txtEnter_Change()
MsfBill.Text = txtEnter.Text
End Sub

Private Sub txtEnter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If MsfBill.Col = 1 Then
     MsfBill.Col = 2
     ArrangeTextbox txtEnter
  ElseIf MsfBill.Col = 2 Then
     MsfBill.Col = 3
     ArrangeTextbox txtEnter
  ElseIf MsfBill.Col = 3 Then
     MsfBill.Col = 4
     ArrangeTextbox txtEnter
  ElseIf MsfBill.Col = 4 Then
      MsfBill.TextMatrix(Indx, 5) = Val(MsfBill.TextMatrix(Indx, 3)) * Val(MsfBill.TextMatrix(Indx, 4))
      FlexgridTotal
      If MsgBox("Do you want to add Additional Items", vbQuestion + vbYesNo + vbDefaultButton1, "Additional Items") = vbYes Then
           MsfBill.Rows = MsfBill.Rows + 1
           Indx = Indx + 1
           MsfBill.Col = 1
           MsfBill.Row = Indx
           MsfBill.TextMatrix(Indx, 0) = Indx
           txtEnter.Visible = False
           ArrangeTextbox cmbItmcode
      Else
          ImgSave_Click
  End If
End If
End If
End Sub
Private Sub FlexgridTotal()
Dim sTot
If Indx = 1 Then
sTot = Val(MsfBill.TextMatrix(Indx, 5))
End If
sTot = Val(txtTotal) + Val(MsfBill.TextMatrix(Indx, 5))
txtTotal.Text = sTot
End Sub
Private Function CalculateTotAmount()
 Dim ToTamt
        ToTamt = 0
         For Inti = 1 To MsfBill.Rows - 1
            ToTamt = ToTamt + Val(MsfBill.TextMatrix(Inti, 3))
        Next
        CalculateTotAmount = Val(ToTamt)
End Function

