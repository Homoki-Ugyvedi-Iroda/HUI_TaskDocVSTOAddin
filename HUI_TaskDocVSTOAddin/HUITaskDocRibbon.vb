Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.SharePoint.Client
Imports SPHelper.SPHUI

Public Class HUITaskDocRibbon
    Private WithEvents mySentItem As Outlook.Items
    Private LastItemFiledGuid As String
    Const DeferredDeliveryTimeinSeconds As Integer = 30
    Public LastTaskFiled As String
    Private MyTimer As New System.Timers.Timer((DeferredDeliveryTimeinSeconds + 20) * 1000)


    Private Async Sub btnReplyToThisEmailFromTask_Click(sender As Object, e As RibbonControlEventArgs) Handles btnReplyToThisEmailFromTask.Click
        Dim MySPHelper = Globals.ThisAddIn.Connection
        Dim Completed As Boolean = cbxWithCompleted.Checked
        Dim SelectedMail As Outlook.MailItem = Globals.ThisAddIn.FormRegionUsed.CurrentMail 'GetSelectedMailItem(Globals.ThisAddIn.Application)
        Dim Results As List(Of TaskClass) = MySPHelper.Tasks.Where(Function(x) x.ConversationID = SelectedMail.ConversationID).ToList
        Results.AddRange(GetTaskClassesFromMailBody(SelectedMail))
        Dim ChosenTask As Integer
        If Results.Count = 0 Then
            btnReplyFromTask_Click(sender, e)
            Exit Sub
        End If
        If Results.Count > 2 Then
            ChosenTask = ChooseFromLimitedTask(Results)
            If ChosenTask = 0 Then Exit Sub
        Else
            ChosenTask = Results.First.ID
        End If
        Dim SelectedTaskListItem As ListItem = Await TaskClass.GetTaskDataFromSPAsync(MySPHelper.Context, ChosenTask)
        ReplyToMailItemBasedonTask(SelectedTaskListItem, SelectedMail, Completed)
    End Sub
    Private Function ChooseFromLimitedTask(tasksToChooseFrom As IEnumerable(Of TaskClass)) As Integer
        Dim TaskChooser As New SPTaxonomy_wWForms_TreeView.SPTreeview
        Dim MySPTaxonomyTreeView As New SPTaxonomy_wWForms_TreeView.SPTreeview
        Dim TaskTreeView = MySPTaxonomyTreeView.ShowNodeswithParents(tasksToChooseFrom, True)
        TaskChooser.CopyTreeNodes(TaskTreeView, TaskChooser.TreeViewBase)
        TaskChooser.ShowDialog()
        If Not TaskChooser.ShowDialog = DialogResult.OK Then Return 0
        Return TaskChooser.SelectedTaskID
    End Function

    Private Async Sub btnReplyFromTask_Click(sender As Object, e As RibbonControlEventArgs) Handles btnReplyFromTask.Click
        Dim TaskChooser As New SPTaxonomy_wWForms_TreeView.SPTreeview
        Dim MySPTaxonomyTreeView As New SPTaxonomy_wWForms_TreeView.SPTreeview
        Dim TaskTreeView = MySPTaxonomyTreeView.ShowNodeswithParents(Globals.ThisAddIn.Connection.Tasks, True)
        TaskChooser.CopyTreeNodes(TaskTreeView, TaskChooser.TreeViewBase)
        '        AddHandler TaskChooser.ShowTaskSelected, AddressOf CreateNewMailItemfromSelectedTask
        TaskChooser.ShowDialog()
        If Not TaskChooser.ShowDialog = DialogResult.OK Then Exit Sub
        Dim Completed As Boolean = cbxWithCompleted.Checked
        Dim SelectedTaskListItem As ListItem = Await TaskClass.GetTaskDataFromSPAsync(Globals.ThisAddIn.Connection.Context, TaskChooser.SelectedTaskID)
        CreateNewMailItemfromTask(SelectedTaskListItem, Completed)
    End Sub
End Class
