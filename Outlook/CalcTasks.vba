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
    End If
      
Wend
    MsgBox "Total time for tasks is " & totalduration & "Min or " & totalduration / 60 & "Hr"
End Sub
