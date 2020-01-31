Imports System.Windows.Forms
Imports SPHelper.SPHUI

Public Class HUI_TaskDocFormRegion

#Region "Form Region Factory"

    <Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)>
    <Microsoft.Office.Tools.Outlook.FormRegionName("HUI_TaskDocVSTOAddin.HUI_TaskDocFormRegion")>
    Partial Public Class HUI_TaskDocFormRegionFactory

        ' Occurs before the form region is initialized.
        ' To prevent the form region from appearing, set e.Cancel to true.
        ' Use e.OutlookItem to get a reference to the current Outlook item.
        Private Sub HUI_TaskDocFormRegionFactory_FormRegionInitializing(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs) Handles Me.FormRegionInitializing

        End Sub

    End Class

#End Region
    Property LastTabPage As Integer
    Property ShowAllPartners As Boolean
    Property TaskTreeView As TreeView
    Property CurrentMail As Outlook.MailItem

    'Occurs before the form region is displayed. 
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub HUI_TaskDocFormRegion_FormRegionShowing(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionShowing
        'Dim NotFound = "Nem találtam előzményt."
        'CurrentMail = GetSelectedMailItem(Globals.ThisAddIn.Application)
        'cbTasksFound.Text = NotFound
        'If Not IsNothing(CurrentMail) AndAlso Not IsNothing(CurrentMail.ConversationID) Then
        '    Dim ResultFromConversationID As List(Of TaskClass) = Globals.ThisAddIn.Connection.Tasks.Where(Function(x) x.ConversationID = CurrentMail.ConversationID).ToList
        '    Dim ResultFromBody As List(Of TaskClass) = GetTaskClassesFromMailBody(CurrentMail)
        '    cbTasksFound.DisplayMember = "TitleOrTaskName"
        '    cbTasksFound.ValueMember = "ID"
        '    cbTasksFound.Items.AddRange(ResultFromConversationID.ToArray)
        '    cbTasksFound.Items.AddRange(ResultFromBody.ToArray)
        '    If cbTasksFound.Items.Count = 0 Then cbTasksFound.Text = NotFound
        'End If
    End Sub


    'Occurs when the form region is closed.   
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub HUI_TaskDocFormRegion_FormRegionClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionClosed

    End Sub

    Private Sub cbFuzzLeElozmenyhez_Click(sender As Object, e As EventArgs)
        If IsNothing(CurrentMail.ConversationID) Then Exit Sub
        '#a cbTasksFound szerinti Selected TaskID-hoz fűzze le ezt az emailt, azaz: 
        '   (A) csatolmányként lementi Taskhoz (egyszerű, de nem lesz kereshető), é/v 
        '   (B) lementi a Task adatai szerinti Dochelyre [oda is rögzíti a TaskID-t és a TaskID egyéb beállításait], ÉS belinkeli a lementést a Taskhoz,
        '       mentési fájlnév és adatok automatizáltan vagy külön ablak [=régi FileMail]? ilyenkor egész emailt menti, csatolmányokat nem kezeli külön

        'Dim result As List(Of TaskClass) = Globals.ThisAddIn.Connection.Tasks.Where(Function(x) x.ConversationID = CurrentMail.ConversationID).ToList
        'Dim TaskTitlesFound As String = String.Empty
        'For Each _task As TaskClass In result
        '    TaskTitlesFound += _task.TitleOrTaskName
        'Next
        'If result.Count = 0 Then lblConversationID.Text = "Nem találtam" Else lblConversationID.Text = TaskTitlesFound
        'Result.Count szerint válogathatunk irányadó Task közül => lefűz ide, lefűz alá új subtaskként [új adatok?], lefűz új taskként [új adatok?], tasklist megnyitása

    End Sub
    'Private Sub lblConversationIDFrissítés()
    '    lblConversationID.Text = "Nem tartozik ehhez a ConversationID-hez Task"
    '    If Not IsNothing(CurrentMail) Then
    '        lblConversationID.Text = "Van mail, de nincsen ConversationID"
    '        If Not IsNothing(CurrentMail.ConversationID) Then lblConversationID.Text = CurrentMail.ConversationID
    '    Else
    '        lblConversationID.Text = "Nincsen mail"
    '    End If
    'End Sub

    Private Sub cbTaskForThisConversationId_SelectedIndexChanged(sender As Object, e As EventArgs)
    End Sub

    Private Sub btnCsakCsatolmanyokLementese_Click(sender As Object, e As EventArgs)
        '#Csativálasztó listával, mentési fájlnév és adatok automatizáltan vagy külön ablak?  
        'legtöbb művelet ezt igényli, nem kizárólag feladathoz mentjük az emailek többségét;
        'esetleg fájllefűzésből fűzzük hozzá mindig taskhoz?
        'előzménykereső ablak?
        'metaadatkitöltés - taskból; -fi
    End Sub

    Private Sub btnChooseOtherTask_Click(sender As Object, e As EventArgs)
        '#Feladatválasztó ablak után lényegében azt kínálja fel, mint btnFuzzLeElozmenyhez esetén
    End Sub

    Private Sub btnNewTask_Click(sender As Object, e As EventArgs)
        '# létrehoz egy új üres taskot, amihez csatolja ezt az emailt/csatolmányt és annak megfelelő, SP szerinti új task ablak megnyitása? 
    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanelTop.Paint

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub cbFolderTypeToUse_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbFolderTypeToUse.SelectedIndexChanged

    End Sub

    Private Sub cbTaskChosenHistoryFileTask_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbTaskChosenHistoryFileTask.SelectedIndexChanged

    End Sub
End Class
