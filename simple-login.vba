Private Sub btnEnter_Click()
  Dim QueriedUser As Recordset, SQL As String, Flag As Boolean
  '****** To declare Database, Table and open the database once instead of repeating process in other Procedures
  Set MyDb = DBEngine.Workspaces(0).Databases(0)
  '******************************************************

  SQL = "SELECT * FROM USERTABLE WHERE ID = '" & txtUserID & "'"
  Flag = False
  Do Until Flag = True
      Set QueriedUser = MyDb.OpenRecordset(SQL)
        If IsNull(txtUserID) Then
          Flag = False
          MsgBox "You must enter User ID in order to log in!"
          Flag = True
          txtUserID.SetFocus
          DoCmd.Hourglass False
        ElseIf QueriedUser.EOF Or QueriedUser.BOF Then
          Flag = False
          MsgBox "The User ID is invalid. Please try again!"
          Flag = True
          txtUserID.SetFocus
        ElseIf IsNull(txtPassword) Then
          Flag = False
          MsgBox "Please enter your password!"
          Flag = True
          txtPassword.SetFocus
          DoCmd.Hourglass False
        ElseIf [QueriedUser]![Password] <> txtPassword Then
          Flag = False
          MsgBox "Wrong Password!"
          Flag = True
          txtPassword.SetFocus
        Else
          Flag = True
          DoCmd.OpenForm "WelcomePage"
  Loop
  QueriedUser.Close
End Sub

Private Sub txtPassword_LostFocus()
    btnEnter_Click
End Sub


    
