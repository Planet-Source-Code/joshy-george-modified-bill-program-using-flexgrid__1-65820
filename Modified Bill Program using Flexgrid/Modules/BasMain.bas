Attribute VB_Name = "BasMain"
'---------> Created by Joshy George <----------
'---------> Use this as you like, no copyright, no restrictions <----------
'---------> All I want is to Give, because I have Received  <----------
'---------> So much help from others.  Thank You ! <----------
'---------> Email:-  joshy_geo@hotmail.com <----------
'---------> WebSite:- www.joshygeo.tk <----------
'---------> if u have any problem in working this project plz contact me through email......>
'Please read
'----------------------
'This Program can Run in Sql server & MS Access
'1--Sql server
'2--MS Access
'If u Select Sql server then Plz Create Database name (SalesBill)
'run Sql Script( I Enclose this file in Sqlserver folder)
'then Create DSN for sql server Name (SalesBill)
'Then add some value in ItemMaster Table direct from SQL server or  Run project then select Itemmaster ......then u can  add items
'--------------------------------------------------------
'if u select MS Access then
'nothing 2 do
'---------------------------------------------------------------------
Option Explicit
Public myConection As New ADODB.Connection
Public Rs As New ADODB.Recordset
Dim RsNewNo As New ADODB.Recordset
Dim Inti, I
Private Sub Main()
'This code for Sql server Connection
'-----------------------------------------------------------------------
'myConection.ConnectionString = "DSN=SalesBill"
'-----------------------------------------------------------------------
'This Code for Access Connection
'-----------------------------------------------------------------------
 myConection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\SalesBill.mdb;Persist Security Info=False"
'-----------------------------------------------------------------------
 myConection.Open
   If myConection.State = adStateOpen Then
       frmMaster.Show
    Else
       MsgBox "Error in Connecting Database please check Connection", vbCritical
   End If
End Sub
'Function is Using for Fill ComboBox
Public Function FillCombo(Cmb As ComboBox, strSQl As String)
On Error GoTo Err_Handler
If Rs.State = 1 Then Rs.Close
Rs.Open strSQl, myConection, adOpenKeyset, adLockOptimistic
If Not Rs.EOF Then
    Cmb.Clear
    Rs.MoveFirst
    Do While Not Rs.EOF
        With Rs
            Cmb.AddItem .Fields(0)
        End With
    Rs.MoveNext
    Loop
End If
Exit Function:
Err_Handler:
If Err.Number > 0 Then
  MsgBox Err.Description, vbCritical
End If
End Function
'This Function for get New number frm database
Public Function GetNewNo(Sql As String) As Long
On Error GoTo Err_Handler
    Dim ID As Long
    Set RsNewNo = myConection.Execute(Sql)
    If Not IsNull(RsNewNo(0)) Then
        ID = RsNewNo(0).Value
    Else
        ID = 10000
    End If
    GetNewNo = ID
    RsNewNo.Close
    Set RsNewNo = Nothing
Exit Function
Err_Handler:
If Err.Number > 0 Then
  MsgBox Err.Description, vbCritical
End If
End Function
' This procedure for Clearing forms ( All textboxes & combo boxes)
Public Sub Clear(frm As Form)
   Dim c As Control
      For Each c In frm.Controls
          If TypeOf c Is TextBox Or TypeOf c Is ComboBox Then
             c.Text = ""
          End If
      Next c
End Sub
' This Function for using Insert Records
Public Function Update1(Table1 As Variant, ParamArray arr() As Variant)
On Error GoTo Err_Handler
    If Rs.State = 1 Then Rs.Close
        Rs.Open "select * from " & Table1, myConection, 3, 3
            Rs.AddNew
                 For I = 0 To UBound(arr())
                     Rs.Fields(I) = arr(I)
                 Next
            Rs.Update
        Rs.Close
Exit Function
Err_Handler:
If Err.Number > 0 Then
  MsgBox Err.Description, vbCritical
End If
End Function
'Procedure is using unLock All TextBoxes
Public Sub TextUnlock(fname As Variant)
Dim txt As Control
Dim lis As Control
     For Each txt In fname
        If TypeOf txt Is TextBox Then txt.Enabled = True
     Next
     For Each lis In fname
        If TypeOf lis Is ComboBox Then lis.Enabled = True
     Next
End Sub

'Procedure is using Lock All TextBoxes

Public Sub TextLock(fname As Variant)
  Dim txt As Control
  Dim lis As Control
        For Each txt In fname
           If TypeOf txt Is TextBox Then txt.Enabled = False
        Next
        For Each lis In fname
           If TypeOf lis Is ComboBox Then lis.Enabled = False
        Next
End Sub
'Function is using to Checking Entering KeyAscii is Numeric or Not
Public Function CheckNumeric(KeyAscii As Integer) As Integer
       CheckNumeric = KeyAscii
       If KeyAscii <> 13 And KeyAscii <> 27 Then
          If KeyAscii < 47 Or KeyAscii > 57 Then
             If KeyAscii <> 8 Then
                CheckNumeric = 0
             End If
          End If
       End If
End Function
'Function is using to Checking Entering KeyAscii is Character or Not
Public Function CheckCharecter(KeyAscii As Integer) As Integer
       CheckCharecter = KeyAscii
       If KeyAscii = 39 Then
          CheckCharecter = 0
       End If
End Function
'Set Selection on focusing TextBox
Public Sub FocusControl(ActiveControl As Control)
       ActiveControl.SelStart = 0
       ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub
'Checking all required fields is Entered
Public Function CheckRequiredFlds(FormName As Form, ReqrdControl As Variant) As Integer
       CheckRequiredFlds = -1
       For Inti = 0 To UBound(ReqrdControl)
           If Len(Trim(FormName.Controls(ReqrdControl(Inti)).Text)) = 0 Then
              CheckRequiredFlds = Inti
              Exit For
           ElseIf Trim(FormName.Controls(ReqrdControl(Inti)).Text) = 0 Then
              CheckRequiredFlds = Inti
              Exit For
           End If
       Next
       If Val(CheckRequiredFlds) >= 0 Then
          MsgBox "Required fields is empty,Can't save Data", vbInformation, "Check Fields"
          FormName.Controls(ReqrdControl(Inti)).SetFocus
       End If
End Function
