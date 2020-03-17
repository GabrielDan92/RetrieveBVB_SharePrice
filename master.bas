Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Sub sharePrice()

    Dim ie As New InternetExplorer, _
    link As String, _
    pret As String, _
    column As String, _
    rng As Range, _
    symbolArr(3) As Variant, _
    lastRow As Long
    
    link = "http://www.bvb.ro/FinancialInstruments/Details/FinancialInstrumentsDetails.aspx?s="

    symbolArr(0) = "BRD"
    symbolArr(1) = "TLV"
    symbolArr(2) = "SNG"
    symbolArr(3) = "SNN"

    For i = 0 To UBound(symbolArr)

            If symbolArr(i) = "BRD" Then
                column = "B"
            ElseIf symbolArr(i) = "TLV" Then
                column = "C"
            ElseIf symbolArr(i) = "SNG" Then
                column = "D"
            ElseIf symbolArr(i) = "SNN" Then
                column = "E"
            End If

            ie.navigate (link & symbolArr(i))
            ie.Visible = True
            Call ieBusy(ie)

            Do While ie.document.querySelector(".value") Is Nothing
                DoEvents
                Sleep 500
                secondsCounter = secondsCounter + 1
                If secondsCounter = 40 Then
                    Exit Do
                End If
            Loop

        pret = ie.document.querySelector(".value").innerHTML

        With ThisWorkbook.Sheets("Sheet1")
            Set rng = .Range("B2").CurrentRegion
            lastRow = rng.Rows.Count + 1
            'if the cell where the price will be added is not empty, add it on the next row
                If .Range(column & lastRow).Value = "" Then
                    If .Range("A" & lastRow).Value = "" Then .Range("A" & lastRow).Value = Now()
                    .Range(column & lastRow).Value = pret
                Else
                    If .Range(column & lastRow + 1).Value = "" Then
                        'if the date is not added, add it
                        If .Range("A" & lastRow + 1).Value = "" Then .Range("A" & lastRow + 1).Value = Now()
                        .Range(column & lastRow + 1).Value = pret
                    Else
                        If .Range("A" & lastRow + 2).Value = "" Then .Range("A" & lastRow + 2).Value = Now()
                        .Range(column & lastRow + 2).Value = pret
                    End If
                End If
        End With
        ThisWorkbook.RefreshAll

    Next i

ie.Quit
Set ie = Nothing

ThisWorkbook.RefreshAll
Set rng = ThisWorkbook.Sheets("Sheet1").Range("B2").CurrentRegion
lastRow = rng.Rows.Count + 1

'==========================sending the emails==================================

    Dim outlookApp As New Outlook.Application
    Dim newEmail As Outlook.MailItem
    Dim ChartObject As ChartObject
    Dim myInspector As Outlook.Inspector
    Dim wdDoc As Word.document
    Dim oWdRng As Word.Range
    
    'reference to chart
        Set ChartObject = ThisWorkbook.Sheets("Sheet1").ChartObjects(1)
        ChartObject.Chart.ChartArea.Copy
    
    'create the table
        htmlContent = "<table border=1>"
            'header row (bold)
                htmlContent = htmlContent & "<tr>"
                    htmlContent = htmlContent & "<th align=center>" & "BRD" & "</th>"
                    htmlContent = htmlContent & "<th align=center>" & "TLV" & "</th>"
                    htmlContent = htmlContent & "<th align=center>" & "SNG" & "</th>"
                    htmlContent = htmlContent & "<th align=center>" & "SNN" & "</th>"
                htmlContent = htmlContent & "</tr>"
            'content row
                htmlContent = htmlContent & "<tr>"
                
                With ThisWorkbook.Sheets("Sheet1")
                    htmlContent = htmlContent & "<td align=center>" & .Range("B" & lastRow).Value & "</td>"
                    htmlContent = htmlContent & "<td align=center>" & .Range("C" & lastRow).Value & "</td>"
                    htmlContent = htmlContent & "<td align=center>" & .Range("D" & lastRow).Value & "</td>"
                    htmlContent = htmlContent & "<td align=center>" & .Range("E" & lastRow).Value & "</td>"
                End With
                
                htmlContent = htmlContent & "</tr>"
                
        htmlContent = htmlContent & "</table>"
    

    'create the email
        Set newEmail = outlookApp.CreateItem(olMailItem)
        newEmail.To = "dan-gabriel.pintoiu@email.com"
        newEmail.Subject = "Stock prices on " & DateValue(Now())
        
    
    'email's body
        newEmail.HTMLBody = "Hello Gabriel, <br><p>" _
        & "Please find below today's stock market prices for your portofolio's symbols: <br><p>" _
        & htmlContent _
        & vbNewLine _
        & "The evolution of the symbols is as follows: <br><p>" _
        & "<b>BRD:</b> " & Format(((ThisWorkbook.Sheets("Sheet1").Range("B" & lastRow).Value - ThisWorkbook.Sheets("Sheet1").Range("B" & lastRow - 1).Value) / ThisWorkbook.Sheets("Sheet1").Range("B" & lastRow - 1).Value), "Percent") & "<br>" _
        & "<b>Banca Transilvania:</b> " & Format(((ThisWorkbook.Sheets("Sheet1").Range("C" & lastRow).Value - ThisWorkbook.Sheets("Sheet1").Range("C" & lastRow - 1).Value) / ThisWorkbook.Sheets("Sheet1").Range("C" & lastRow - 1).Value), "Percent") & "<br>" _
        & "<b>Romgaz:</b> " & Format(((ThisWorkbook.Sheets("Sheet1").Range("D" & lastRow).Value - ThisWorkbook.Sheets("Sheet1").Range("D" & lastRow - 1).Value) / ThisWorkbook.Sheets("Sheet1").Range("D" & lastRow - 1).Value), "Percent") & "<br>" _
        & "<b>Nuclearelectrica:</b> " & Format(((ThisWorkbook.Sheets("Sheet1").Range("E" & lastRow).Value - ThisWorkbook.Sheets("Sheet1").Range("E" & lastRow - 1).Value) / ThisWorkbook.Sheets("Sheet1").Range("E" & lastRow - 1).Value), "Percent") & "<br>" _
        & " <br><p>"
    
    'see the email
        newEmail.display
        
    'Get the Word Editor
        Set wdDoc = newEmail.GetInspector.WordEditor

    'Define the range, insert a blank line, collapse the selection.
        Set oWdRng = wdDoc.Application.ActiveDocument.Content
            oWdRng.InsertAfter " " & vbNewLine
            oWdRng.collapse Direction:=wdCollapseEnd
                
    'Paste the object.
        oWdRng.Paste

    'send the email
        newEmail.send
        
End Sub

Sub ieBusy(ie As Object)
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop
End Sub
