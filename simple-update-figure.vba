Private Sub UpdateFigures()
    Dim db As Database, QRY As QueryDef, ChosenMajorLine As String, ChosenMonth As String
    Dim Quarter1 As Variant, Quarter2 As Variant, Quarter3 As Variant, Quarter4 As Variant, i As Integer
    Dim strArrQuarter As String
    Set db = CurrentDb()
    DoCmd.SetWarnings False
    
    ChosenMajorLine = Me.dlistMajorLine.Value
    ChosenMonth = Me.listMonth.Value
    
    If ChosenMonth = "01" Or ChosenMonth = "02" Or ChosenMonth = "03" Then
        strArrQuarter = "('01','02','03')"
    ElseIf ChosenMonth = "04" Or ChosenMonth = "05" Or ChosenMonth = "06" Then
        strArrQuarter = "('04','05','06')"
    ElseIf ChosenMonth = "07" Or ChosenMonth = "08" Or ChosenMonth = "09" Then
        strArrQuarter = "('07','08','09')"
    ElseIf ChosenMonth = "10" Or ChosenMonth = "11" Or ChosenMonth = "12" Then
        strArrQuarter = "('10','11','12')"
    Else
        MsgBox "Couldn't detect the Month: " & ChosenMonth
    End If
    
    For Each QRY In db.QueryDefs
        If QRY.Name = "ForMajorlineByMonth" Then
            QRY.SQL = "SELECT * FROM MajorLineGPW WHERE Majorline = '" & ChosenMajorLine & "';"
        End If
    Next QRY
    
    Me.txtMTDActual.RowSource = "SELECT CCur(TotalPremium) AS Total FROM ForMajorlineByMonth WHERE Class = 'Actual' AND FinancialMonth = '" & ChosenMonth & "';"
    Me.txtQTDActual.RowSource = "SELECT CCur(Sum(TotalPremium)) AS Total FROM ForMajorlineByMonth WHERE Class='Actual' AND FinancialMonth IN " & strArrQuarter & ";"
    Me.txtYTDActual.RowSource = "SELECT CCur(Sum(TotalPremium)) AS Total FROM ForMajorlineByMonth WHERE Class='Actual';"
    
    Me.txtMTDvsBudget.RowSource = "SELECT FormatPercent((A.Actual / B.Budget)-1) AS vsBudget FROM (SELECT TotalPremium AS Budget, * FROM ForMajorlineByMonth WHERE Class = 'Budget')  AS B LEFT JOIN (SELECT TotalPremium AS Actual, * FROM ForMajorlineByMonth WHERE Class = 'Actual')  AS A ON (A.FinancialMonth = B.FinancialMonth) AND (A.Majorline = B.Majorline) WHERE B.FinancialMonth = '" & ChosenMonth & "';"
    Me.txtQTDvsBudget.RowSource = "SELECT FormatPercent(Sum(A.Actual)/Sum(B.Budget)-1) AS vsBudget FROM (SELECT TotalPremium AS Budget, * FROM ForMajorlineByMonth WHERE Class = 'Budget')  AS B LEFT JOIN (SELECT TotalPremium AS Actual, * FROM ForMajorlineByMonth WHERE Class = 'Actual')  AS A ON (B.Majorline = A.Majorline) AND (B.FinancialMonth = A.FinancialMonth) WHERE B.FinancialMonth IN " & strArrQuarter & " GROUP BY B.Majorline;"
    Me.txtYTDvsBudget.RowSource = "SELECT FormatPercent(Sum(A.Actual)/Sum(B.Budget)-1) AS vsBudget FROM (SELECT TotalPremium AS Budget, * FROM ForMajorlineByMonth WHERE Class = 'Budget')  AS B LEFT JOIN (SELECT TotalPremium AS Actual, * FROM ForMajorlineByMonth WHERE Class = 'Actual')  AS A ON (B.FinancialMonth = A.FinancialMonth) AND (B.Majorline = A.Majorline) GROUP BY B.Majorline;"
    
    Me.txtMTDvsPrior.RowSource = "SELECT FormatPercent((A.Actual/P.Prior)-1) AS vsPrior FROM (SELECT TotalPremium AS Actual, * FROM ForMajorlineByMonth WHERE Class = 'Actual')  AS A LEFT JOIN (SELECT TotalPremium AS [Prior], * FROM ForMajorlineByMonth WHERE Class = 'Prior')  AS P ON (A.Majorline = P.Majorline) AND (A.FinancialMonth = P.FinancialMonth) WHERE A.FinancialMonth = '" & ChosenMonth & "';"
    Me.txtQTDvsPrior.RowSource = "SELECT FormatPercent(Sum(A.Actual)/Sum(P.Prior)-1) AS vsPrior FROM (SELECT TotalPremium AS Actual, * FROM ForMajorlineByMonth WHERE Class = 'Actual')  AS A LEFT JOIN (SELECT TotalPremium AS [Prior], * FROM ForMajorlineByMonth WHERE Class = 'Prior')  AS P ON (A.Majorline = P.Majorline) AND (A.FinancialMonth = P.FinancialMonth) WHERE A.FinancialMonth IN " & strArrQuarter & " GROUP BY A.Majorline;"
    Me.txtYTDvsPrior.RowSource = "SELECT FormatPercent(Sum(A.Actual)/Sum(P.Prior)-1) AS vsPrior FROM (SELECT TotalPremium AS Actual, * FROM ForMajorlineByMonth WHERE Class = 'Actual')  AS A LEFT JOIN (SELECT TotalPremium AS [Prior], * FROM ForMajorlineByMonth WHERE Class = 'Prior')  AS P ON (A.Majorline = P.Majorline) AND (A.FinancialMonth = P.FinancialMonth) GROUP BY A.Majorline;"
    
    Me.txtMTDActual.Requery
    Me.txtMTDvsBudget.Requery
    Me.txtMTDvsPrior.Requery
    Me.txtQTDActual.Requery
    Me.txtQTDvsBudget.Requery
    Me.txtQTDvsPrior.Requery
    Me.txtYTDActual.Requery
    Me.txtYTDvsBudget.Requery
    Me.txtYTDvsPrior.Requery
        
    Me.chartMajorLine.Requery
    Me.chartMajorLine.ChartTitle = ChosenMajorLine & " - GPW By Month"
    
    DoCmd.SetWarnings True
End Sub
'*******************************************************************************************************************
Private Sub dlistMajorLine_Change()
    UpdateFigures
End Sub
'*******************************************************************************************************************
Private Sub listMonth_Change()
    UpdateFigures
End Sub
