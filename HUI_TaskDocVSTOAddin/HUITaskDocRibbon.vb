Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.SharePoint.Client

Public Class HUITaskDocRibbon
    Private WithEvents mySentItem As Outlook.Items
    Private LastItemFiledGuid As String
    Const DeferredDeliveryTimeinSeconds As Integer = 30
    Public LastTaskFiled As String
    Private MyTimer As New System.Timers.Timer((DeferredDeliveryTimeinSeconds + 20) * 1000)

    Private Sub HUITaskDocRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnReplyToThisEmailFromTask_Click(sender As Object, e As RibbonControlEventArgs) Handles btnReplyToThisEmailFromTask.Click
        Dim Completed As Boolean = cbxWithCompleted.Checked
        Dim SelectedTaskListItem As ListItem
        Dim SelectedMail As Outlook.MailItem
        '#kidolgozás fenti kettő sor választásnak
        ReplyToMailItemBasedonTask(SelectedTaskListItem, SelectedMail, Completed)
    End Sub

    Private Sub btnReplyFromTask_Click(sender As Object, e As RibbonControlEventArgs) Handles btnReplyFromTask.Click
        Dim Completed As Boolean = cbxWithCompleted.Checked
        Dim SelectedTaskListItem As ListItem
        '#kidolgozás fenti kettő sor választásnak
        CreateNewMailItemfromTask(SelectedTaskListItem, Completed)

    End Sub
End Class
