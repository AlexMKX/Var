Public Sub CalcTime()
 Dim objView As View
 Dim tableView As Outlook.tableView
Set tableView = Application.ActiveExplorer.CurrentView
'If tableView.Name <> "Tasks for today" Then Exit Sub
Dim table As Outlook.table
Set table = tableView.GetTable

Set NS = Application.GetNamespace("MAPI")
'NS.Logon

Dim totalduration As Integer
totalduration = 0
Dim smalltasks, bigtasks As Integer
smalltasks = 0
bigtasks = 0
While Not table.EndOfTable
    Dim Entry As String
    
    Dim nextRow As Outlook.Row
    Set nextRow = table.GetNextRow
        
    EntryID = nextRow.Item(1)
    Set Msg = NS.GetItemFromID(EntryID)
    Dim prop As UserProperty
    Set prop = Msg.ItemProperties.Item("Estimation")
    If Not prop Is Nothing Then
        Dim duration As Integer
        duration = prop.Value
        totalduration = totalduration + duration
        If duration > 15 Then bigtasks = bigtasks + 1 Else smalltasks = smalltasks + 1
    End If
      
Wend
Dim switchtime, totaltime As Integer
switchtime = 15 * bigtasks + 5 * smalltasks
totaltime = switchtime + totalduration
    MsgBox "Total time for tasks is " & totalduration & "Min or " & totalduration / 60 & "Hr" & vbCrLf & _
    bigtasks & " big tasks. " & smalltasks & " small tasks. Switch time " & switchtime & "min." & vbCrLf & _
    "Total time spending for today is " & totaltime & "min or " & totaltime / 60 & "H"
End Sub
