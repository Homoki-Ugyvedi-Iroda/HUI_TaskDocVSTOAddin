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
        'Feltölti a három cb-t értékekkel, mindhárom cb értéke azonos, csak a Taskból és a body-ból veszi ki a history-t
        LoadValuesIntocbTaskChosen({cbTaskChosenHistoryNewTask, cbTaskChosenHistoryFileTask, cbFileHistory})
        cbFolderTypeToUse.SelectedValue = "Default"
        LoadValuesIncbMattersAllActive()
        LoadValuesIncbPartnersAllActive()
    End Sub


    'Occurs when the form region is closed.   
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub HUI_TaskDocFormRegion_FormRegionClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionClosed

    End Sub
    Private Sub LoadValuesIntocbTaskChosen(cbControls As IEnumerable(Of ComboBox))
        CurrentMail = GetSelectedMailItem(Globals.ThisAddIn.Application)
        If Not IsNothing(CurrentMail) AndAlso Not IsNothing(CurrentMail.ConversationID) Then
            Dim ResultFromConversationID As List(Of TaskClass) = Globals.ThisAddIn.Connection.Tasks.Where(Function(x) x.ConversationID = CurrentMail.ConversationID).ToList
            Dim ResultFromBody As List(Of TaskClass) = GetTaskClassesFromMailBody(CurrentMail)
            For Each comboBox In cbControls
                comboBox.DisplayMember = "TitleOrTaskName"
                comboBox.ValueMember = "ID"
                comboBox.Items.AddRange(ResultFromConversationID.ToArray)
                comboBox.Items.AddRange(ResultFromBody.ToArray)
            Next
        End If
    End Sub
    Private Sub LoadValuesIncbMattersAllActive()
        cbMatter.DisplayMember = "Value"
        cbMatter.ValueMember = "ID"
        cbMatter.Items.AddRange(Globals.ThisAddIn.Connection.Matters.Where(Function(x) x.Active = True).ToArray)
    End Sub

    Private Sub LoadValuesIncbPartnersAllActive()
        cbPartner.DisplayMember = "Title"
        cbPartner.ValueMember = "ID"
        cbMatter.Items.AddRange(Globals.ThisAddIn.Connection.Persons.Where(Function(x) x.Active = True).ToArray)
        lbTotalPartners.DisplayMember = "Title"
        lbTotalPartners.ValueMember = "ID"
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

    'Private Sub btnCsakCsatolmanyokLementese_Click(sender As Object, e As EventArgs)
    '    '#Csativálasztó listával, mentési fájlnév és adatok automatizáltan vagy külön ablak?  
    '    'legtöbb művelet ezt igényli, nem kizárólag feladathoz mentjük az emailek többségét;
    '    'esetleg fájllefűzésből fűzzük hozzá mindig taskhoz?
    '    'előzménykereső ablak?
    '    'metaadatkitöltés - taskból; -fi
    'End Sub

    Private Sub btnChooseOtherTask_Click(sender As Object, e As EventArgs)
        '#Feladatválasztó ablak után lényegében azt kínálja fel, mint btnFuzzLeElozmenyhez esetén
    End Sub

    Private Sub btnNewTask_Click(sender As Object, e As EventArgs)
        '# létrehoz egy új üres taskot, amihez csatolja ezt az emailt/csatolmányt és annak megfelelő, SP szerinti új task ablak megnyitása? 
    End Sub

    Private Sub btnHistoryChosenAsTemplateForNewTask_Click(sender As Object, e As EventArgs) Handles btnHistoryChosenAsTemplateForNewTask.Click
        If IsNothing(cbTaskChosenHistoryNewTask.SelectedValue) Then Exit Sub
        '#Mentse le egy új taskként a korábbi adataival, majd nyissa meg egy böngészőablakban az új taskot
    End Sub

    Private Sub btnExistingTaskChoiceAsTemplateForNewTask_Click(sender As Object, e As EventArgs) Handles btnExistingTaskChoiceAsTemplateForNewTask.Click
        '#feladatválasztó ablakot hívja meg
    End Sub

    'Private Sub btnFileToChosenHistoryFileTask_Click(sender As Object, e As EventArgs) Handles btnFileToChosenHistoryFileTask.Click
    '    If IsNothing(cbTaskChosenHistoryFileTask.SelectedValue) Then Exit Sub
    '    '#Mentse le az itemet a korábbi task metaadataival, majd fűzze hozzá related item-ként a korábbi taskhoz

    'End Sub

    Private Sub btnFileToDocLibrary_Click(sender As Object, e As EventArgs) Handles btnFileToDocLibrary.Click
        '#Mentse le az itemet a kiválasztott matter és partner metaadatokkal.
        'Ha volt korábbi task kiválasztva a cbFileHistory-ban, akkor fűzze hozzá a taskhoz related item-ként az új itemet

    End Sub

    Private Sub btnExistingTaskChoiceAsFileTo_File_Click(sender As Object, e As EventArgs) Handles btnExistingTaskChoiceAsFileTo_File.Click
        '#feladatválasztó ablakot hívja meg
    End Sub

    Private Sub btnAddPartnerFile_Click(sender As Object, e As EventArgs) Handles btnAddPartnerFile.Click
        If IsNothing(cbPartner.SelectedValue) Then Exit Sub
        lbTotalPartners.Items.Add(cbPartner.SelectedItem)
    End Sub

    Private Sub btnDeleteLastPartner_Click(sender As Object, e As EventArgs) Handles btnDeleteLastPartner.Click
        If lbTotalPartners.Items.Count < 1 Then Exit Sub
        lbTotalPartners.Items.Remove(lbTotalPartners.Items.Count - 1)
    End Sub

    Private Sub btnOpenFolder_Click(sender As Object, e As EventArgs) Handles btnOpenFolder.Click
        '#tbPathtoSaveTo szerinti útvonalat nyissa meg böngészőben
    End Sub

    Private Sub cbFolderTypeToUse_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbFolderTypeToUse.SelectedIndexChanged
        '#Generálja újra az útvonalat a kiválasztott típusnak megfelelően
        'GetPathfromSecondLevelValues(cbMatterFile.SelectedItem, cbWorkDocType:=cbWorkDocType, ProjectorSystemText:=tbProjectSystemFile.Text, cbNonWorkDocType:=cbNonWorkDocType, PartnerTable:=PartnersFile, KeywordTable:=KeywordsFile)

    End Sub

    Private Sub btnPartnerQryMailBody_Click(sender As Object, e As EventArgs) Handles btnPartnerQryMailBody.Click
        Dim PartnersToAdd As New List(Of PersonClass)
        Dim SelectedMatter As MatterClass = cbMatter.SelectedItem
        If IsNothing(SelectedMatter) Then Exit Sub
        PartnersToAdd = GetDefaultNonMatterPartnersfromBodyText(CurrentMail, SelectedMatter)
        Dim DefaultPartner As PersonClass = Globals.ThisAddIn.Connection.Persons.Where(Function(x) x.Id = SelectedMatter.MatterDefaultPerson)
        If Not IsNothing(DefaultPartner) Then
            PartnersToAdd.Add(DefaultPartner)
        End If
        lbTotalPartners.Items.AddRange(PartnersToAdd.ToArray)
    End Sub
End Class
