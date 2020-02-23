Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports SPHelper.SPHUI
Imports SPHelper.SPFileFolder
Imports CSOM = Microsoft.SharePoint.Client


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

    'Occurs when the form region is closed.   
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub HUI_TaskDocFormRegion_FormRegionClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionClosed
        Me.Open = False
    End Sub
    Private Sub LoadValuesIntocbTaskChosen(cbControls As IEnumerable(Of ComboBox), currentMail As MailItem)
        If Not IsNothing(currentMail) AndAlso Not IsNothing(currentMail.ConversationID) Then
            Dim ResultFromConversationID As List(Of TaskClass) = Globals.ThisAddIn.Connection.Tasks.Where(Function(x) x.ConversationID = currentMail.ConversationID).ToList
            Dim ResultFromBody As List(Of TaskClass) = GetTaskClassesFromMailBody(currentMail)
            For Each comboBox In cbControls
                comboBox.DisplayMember = "TitleOrTaskName"
                comboBox.ValueMember = "ID"
                comboBox.Items.AddRange(ResultFromConversationID.ToArray)
                comboBox.Items.AddRange(ResultFromBody.ToArray)
            Next
        End If
    End Sub
    'Occurs before the form region is displayed. 
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub HUI_TaskDocFormRegion_FormRegionShowing(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionShowing
        'Feltölti a három cb-t értékekkel, mindhárom cb értéke azonos, csak a Taskból és a body-ból veszi ki a history-t
        'TabFileAsTask.Visible = False
        'TabFileAsTask.Enabled = False

        Me.Open = True
        CurrentMail = TryCast(Me.OutlookItem, MailItem)
        Globals.ThisAddIn.FormRegionUsed = Me
        If IsNothing(CurrentMail) Then Exit Sub
        LoadValuesIntocbTaskChosen({cbTaskChosenHistoryNewTask, cbFileHistory}, CurrentMail)
        LoadValuesIncbMattersAllActive()
        LoadValuesIncbPartnersAllActive()
        If Not OutlookApplicationHelper.IsNewAndEmpty(CurrentMail) Then
            tbTitleFile.Text = CurrentMail.Subject
            LoadAttachmentNames(CurrentMail)
        End If
    End Sub

#End Region
    Public Enum Tab
        FileToDocLibrary = 0
        CreateNewTask = 1
    End Enum
    Property Open As Boolean
    Property LastTabPage As Integer
    Property ShowAllPartners As Boolean
    Property CurrentMail As Outlook.MailItem
    Enum AttachmentIndex
        FullEmail = 254
        EmailWithoutAttachment = 255
    End Enum
    Class HistoryItem
        Property Matter As MatterClass
        Property Partner As List(Of PersonClass)
        Property Task As TaskClass
    End Class



#Region "CreateNewTask_Tab"
    ''' <summary>
    ''' A history listában lévő taskok közül választ, és ez a kiválasztott task szolgál alapjául a metaadatokat illetően egy új taskhoz, 
    ''' ami új taskhoz hozzáfűzi az aktuális emailt vagy csatolmánytg
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Async Sub btnHistoryChosenAsTemplateForNewTask_Click(sender As Object, e As EventArgs) Handles btnHistoryChosenAsTemplateForNewTask.Click
        If IsNothing(cbTaskChosenHistoryNewTask.SelectedItem) Then Exit Sub
        Dim InterimDoc As DocClass = GetDocClassFromFileTab(ReturnSavedFileName())
        Dim FileLeafTarget As String = CheckUnicityFileNameInFolder(Globals.ThisAddIn.Connection.Context, InterimDoc.PathtoSaveTo, InterimDoc.FileName)
        Await CreateAndFileAsFile(FileLeafTarget, InterimDoc)
        'Mentse le egy új taskként a korábbi adataival, majd nyissa meg egy böngészőablakban az új taskot!
        CreateNewTaskBasedOnOld(ReturnTaskFromHistoryItem(cbTaskChosenHistoryNewTask))
        'If IsTaskSelectedAsFileHistoryItem() Then
        '    'A cbTaskChosenHistory-ban lévő korábbi taskhoz fűzi related item-ként az új itemet
        '    MySPHelper.SetRelatedItemForTaskSecond(MySPHelper.Context, _ListItem, NewTask.ID, FileLeafTarget)
        'End If
    End Sub
    ''' <summary>
    ''' Meglévő taskot kijelöl és utána ez a task szolgál egy új task alapjául a metaadatokat illetően,
    ''' ami új taskhoz hozzáfűzi az aktuális emailt vagy csatolmányt.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Async Sub btnExistingTaskChoiceAsTemplateForNewTask_Click(sender As Object, e As EventArgs) Handles btnExistingTaskChoiceAsTemplateForNewTask.Click
        Dim TaskChooser As New SPTaxonomy_wWForms_TreeView.SPTreeview
        TaskChooser.CopyTreeNodes(Globals.ThisAddIn.TaskTreeView, TaskChooser.TreeViewBase)
        'AddHandler TaskChooser.ShowTaskSelected, AddressOf NewTaskWithChosenTaskMetadata
        If Not TaskChooser.ShowDialog = DialogResult.OK Then Exit Sub
        Dim MySPHelper = Globals.ThisAddIn.Connection
        Dim SourceTask As TaskClass = Globals.ThisAddIn.Connection.Tasks.Where(Function(x) x.ID = TaskChooser.SelectedTaskID)
        Dim InterimDoc As DocClass = GetDocClassFromSelectedTask(SourceTask, ReturnSavedFileName())
        Dim FileLeafTarget As String = CheckUnicityFileNameInFolder(MySPHelper.Context, InterimDoc.PathtoSaveTo, InterimDoc.FileName)
        Dim _ListItem As CSOM.ListItem = Await CreateAndFileAsFile(FileLeafTarget, InterimDoc)
        CreateNewTaskBasedOnOld(SourceTask)
        MySPHelper.SetRelatedItemForTaskSecond(MySPHelper.Context, _ListItem, SourceTask.ID, FileLeafTarget)
        AddToHistoryFromSelectedTaskAndClear(SourceTask)

    End Sub
    '''' <summary>
    '''' A btnExistingTaskChoiceAsTemplateForNewTask_Click által meghívott aszinkron végrehajtó rész
    '''' </summary>
    '''' <param name="taskID"></param>
    'Private Async Sub NewTaskWithChosenTaskMetadata(taskID As Integer)
    '    Dim MySPHelper = Globals.ThisAddIn.Connection
    '    Dim SourceTask As TaskClass = Globals.ThisAddIn.Connection.Tasks.Where(Function(x) x.ID = taskID)
    '    Dim InterimDoc As DocClass = GetDocClassFromSelectedTask(SourceTask, ReturnSavedFileName())
    '    Dim FileLeafTarget As String = CheckUnicityFileNameInFolder(MySPHelper.Context, InterimDoc.PathtoSaveTo, InterimDoc.FileName)
    '    Dim _ListItem As CSOM.ListItem = Await CreateAndFileAsFile(FileLeafTarget, InterimDoc)
    '    CreateNewTaskBasedOnOld(SourceTask)
    '    MySPHelper.SetRelatedItemForTaskSecond(MySPHelper.Context, _ListItem, SourceTask.ID, FileLeafTarget)
    '    AddToHistoryFromSelectedTaskAndClear(SourceTask)
    'End Sub
    ''' <summary>
    ''' btnCreateNewTask_Click által meghívott aszinkron végrehajtó rész
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Async Sub btnCreateNewTask_Click(sender As Object, e As EventArgs) Handles btnCreateNewTask.Click
        'Létrehoz egy új üres taskot, amihez csatolja ezt az emailt/csatolmányt és annak megfelelő, SP szerinti új task ablak megnyitása. 
        Dim MySPHelper = Globals.ThisAddIn.Connection
        Dim InterimDoc As DocClass = GetDocClassFromFileTab(ReturnSavedFileName())
        Dim FileLeafTarget As String = CheckUnicityFileNameInFolder(MySPHelper.Context, InterimDoc.PathtoSaveTo, InterimDoc.FileName)
        Dim _ListItem As CSOM.ListItem = Await CreateAndFileAsFile(FileLeafTarget, InterimDoc)
        Dim NewTask = CreateNewTask()
        MySPHelper.SetRelatedItemForTaskSecond(MySPHelper.Context, _ListItem, NewTask.ID, FileLeafTarget)
        AddToHistoryFromSelectedTaskAndClear(NewTask)
    End Sub
    'Private Async Sub NewTaskWithEmail(taskID As Integer)
    '    Dim MySPHelper = Globals.ThisAddIn.Connection
    '    Dim SourceTask As TaskClass = Globals.ThisAddIn.Connection.Tasks.Where(Function(x) x.ID = taskID)
    '    Dim InterimDoc As DocClass = GetDocClassFromFileTab(ReturnSavedFileName())
    '    Dim FileLeafTarget As String = CheckUnicity(InterimDoc.PathtoSaveTo, InterimDoc.FileName)
    '    Dim _ListItem As CSOM.ListItem = Await CreateAndFileAsFile(FileLeafTarget, InterimDoc)
    '    CreateNewTask()
    '    MySPHelper.SetRelatedItemForTaskSecond(MySPHelper.Context, _ListItem, SourceTask.ID, FileLeafTarget)
    '    AddToHistoryFromSelectedTaskAndClear(SourceTask)
    'End Sub
#End Region
#Region "FileToDocLibrary_Tab"
    ''' <summary>
    ''' A FileToDocLibrary_Tab adatai alapján lementi az emailt vagy csatolmányát, majd nyissa meg a böngészőablakban a lementett új fájlt.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Async Sub btnFileToDocLibrary_Click(sender As Object, e As EventArgs) Handles btnFileToDocLibrary.Click
        Dim InterimDoc As DocClass = GetDocClassFromFileTab(ReturnSavedFileName())
        Dim FileLeafTarget As String = CheckUnicityFileNameInFolder(Globals.ThisAddIn.Connection.Context, InterimDoc.PathtoSaveTo, InterimDoc.FileName)
        Dim _ListItem As CSOM.ListItem = Await CreateAndFileAsFile(FileLeafTarget, InterimDoc)
        If IsTaskSelectedAsFileHistoryItem() Then
            Dim MySPHelper = Globals.ThisAddIn.Connection
            'Ha volt korábbi task kiválasztva a cbFileHistory-ban, akkor fűzze hozzá a taskhoz related item-ként az új itemet
            MySPHelper.SetRelatedItemForTaskSecond(MySPHelper.Context, _ListItem, ReturnTaskFromHistoryItem(cbFileHistory).ID, FileLeafTarget)
            'MySPHelper.SetRelatedItemForTask(MySPHelper.Context, ReturnTaskFromFilehistoryItem.ID, FileLeafTarget)
        End If
        AddToHistoryFromFileAndClear()
        'Nyissa meg egy böngészőablakban az új file-t!
        DocClass.OpenDocToEdit(_ListItem.Id)
    End Sub
#End Region
#Region "TaskDocFormRegion egyéb controlok és eventek kezelése"
    Private Sub LoadValuesIncbMattersAllActive()
        cbMatter.DisplayMember = "Value"
        cbMatter.ValueMember = "ID"
        cbMatter.Items.AddRange(Globals.ThisAddIn.Connection.Matters.Where(Function(x) x.Active = True).ToArray)
        Dim SelectedMatterFromSender = GetMatterfromSenderData(CurrentMail)
        If Not IsNothing(SelectedMatterFromSender) AndAlso SelectedMatterFromSender.Count > 0 Then
            cbMatter.SelectedItem = SelectedMatterFromSender.First
        End If
    End Sub

    Private Sub LoadValuesIncbPartnersAllActive()
        cbPartner.DisplayMember = "Value"
        cbPartner.ValueMember = "ID"
        lbTotalPartners.DisplayMember = "Value"
        lbTotalPartners.ValueMember = "ID"
        cbPartner.Items.AddRange(Globals.ThisAddIn.Connection.Persons.Where(Function(x) x.Active = True).ToArray)
    End Sub

    Private Sub LoadAttachmentNames(currentMail As MailItem)
        Dim AttachmentNames = GetAllAttachmentDisplayNames(currentMail)
        Dim FullEmail As New OutlookItemAttachment With
            {.DisplayName = "The full email (including attachments)", .ID = AttachmentIndex.FullEmail}
        Dim EmailWithoutAttachment As New OutlookItemAttachment With
            {.DisplayName = "The email without attachments", .ID = AttachmentIndex.EmailWithoutAttachment}
        cbFileEmailOrAttachments.DisplayMember = "DisplayName"
        cbFileEmailOrAttachments.ValueMember = "ID"
        cbFileEmailOrAttachments.Items.AddRange({FullEmail, EmailWithoutAttachment})
        cbFileEmailOrAttachments.Items.AddRange(GetAllAttachments(currentMail).ToArray)
        cbFileEmailOrAttachments.SelectedItem = FullEmail
    End Sub
    Private Sub btnAddPartnerFile_Click(sender As Object, e As EventArgs) Handles btnAddPartnerFile.Click
        If IsNothing(cbPartner.SelectedItem) Then Exit Sub
        lbTotalPartners.Items.Add(cbPartner.SelectedItem)
        CheckIfPathChangesForPartnerChange(sender, e)
    End Sub
    Private Sub btnDeleteLastPartner_Click(sender As Object, e As EventArgs) Handles btnDeleteLastPartner.Click
        If lbTotalPartners.Items.Count < 1 Then Exit Sub
        lbTotalPartners.Items.RemoveAt(lbTotalPartners.Items.Count - 1)
        CheckIfPathChangesForPartnerChange(sender, e)
    End Sub
    Private Sub btnChangePartnerOrder_Click(sender As Object, e As EventArgs) Handles btnChangePartnerOrder.Click
        If lbTotalPartners.Items.Count < 2 Then Exit Sub
        Dim InterimItem As Object = lbTotalPartners.Items(0)
        lbTotalPartners.Items.RemoveAt(0)
        lbTotalPartners.Items.Add(InterimItem)
        CheckIfPathChangesForPartnerChange(sender, e)
    End Sub
    Private Sub btnOpenFolder_Click(sender As Object, e As EventArgs) Handles btnOpenFolder.Click
        Try
            Dim SpcInfo As ProcessStartInfo = New ProcessStartInfo(ReturnUrlToOpen)
            Process.Start(SpcInfo)
        Catch ex As System.Exception
        End Try
    End Sub
    Private Sub btnCreateFolderIfNotExisting_Click(sender As Object, e As EventArgs) Handles btnCreateFolderIfNotExisting.Click
        Dim TargetUrl = ReturnUrlToOpen()
        'Ellenőrizze, hogy a tbPathToSaveTo létezik-e, és ha nem, hozza létre
        If SPHelper.SPFileFolder.TryGetFolderByServerRelativeUrl(Globals.ThisAddIn.Connection.Context, TargetUrl) Then Exit Sub
        SPHelper.SPFileFolder.CreateFolder(Globals.ThisAddIn.Connection.Context.Web, DocClass.ClientDocumentLibraryName, TargetUrl)
    End Sub
    Private Sub btnPartnerQryMailBody_Click(sender As Object, e As EventArgs) Handles btnPartnerQryMailBody.Click
        Dim PartnersToAdd As New List(Of PersonClass)
        Dim SelectedMatter As MatterClass = cbMatter.SelectedItem
        If IsNothing(SelectedMatter) Then Exit Sub
        PartnersToAdd = GetDefaultNonMatterPartnersfromBodyText(CurrentMail, SelectedMatter)
        Dim DefaultPartner As PersonClass = Globals.ThisAddIn.Connection.Persons.Where(Function(x) x.Id = SelectedMatter.MatterDefaultPerson).First
        If Not IsNothing(DefaultPartner) Then
            PartnersToAdd.Add(DefaultPartner)
        End If
        lbTotalPartners.Items.AddRange(PartnersToAdd.ToArray)
    End Sub
    Private Sub cbMatter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbMatter.SelectedIndexChanged
        If IsNothing(cbMatter.SelectedItem) Then Exit Sub
        Dim PartnerSelected As PersonClass = Nothing
        If lbTotalPartners.Items.Count > 0 Then PartnerSelected = lbTotalPartners.Items(0)
        Dim NewDirectory = SPHelper.SPHUIFileFolder.ReturnPathFromMatterAndPartner(cbMatter.SelectedItem, PartnerSelected)
        If String.IsNullOrWhiteSpace(NewDirectory) Then Exit Sub
        tbPathToSaveTo.Text = NewDirectory
    End Sub

    Private Sub cbFileEmailOrAttachments_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbFileEmailOrAttachments.SelectedValueChanged
        If cbFileEmailOrAttachments.SelectedIndex < 2 Then Exit Sub
        Dim SelectedAttachment As Attachment = cbFileEmailOrAttachments.SelectedItem
        tbTitleFile.Text = SelectedAttachment.DisplayName
    End Sub

    Private Sub TabCreateTask_Enter(sender As Object, e As EventArgs) Handles TabCreateTask.Enter
        Me.LastTabPage = Tab.CreateNewTask
    End Sub
    Private Sub TabFileAsDoc_Enter(sender As Object, e As EventArgs) Handles TabFileAsDoc.Enter
        Me.LastTabPage = Tab.FileToDocLibrary
    End Sub

    Private Sub cbTaskChosenHistoryNewTask_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbTaskChosenHistoryNewTask.SelectedValueChanged
        btnHistoryChosenAsTemplateForNewTask.Enabled = True
    End Sub

    Private Sub CheckIfPathChangesForPartnerChange(sender As Object, e As EventArgs)
        cbMatter_SelectedIndexChanged(sender, e)
    End Sub
    'Private Sub TabControl1_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl1.Selected
    '    If TabControl1.SelectedTab.Equals(TabFileAsDoc) Then
    '        Me.LastTabPage = Tab.FileToDocLibrary
    '    ElseIf TabControl1.SelectedTab.Equals(TabCreateTask) Then
    '        Me.LastTabPage = Tab.FileToDocLibrary
    '    End If
    'End Sub
#End Region
#Region "Not used FileToTaskAsItem"
    ''' <summary>
    ''' Nem használt függvény FileToTaskAsItem fül alapján
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnTaskChoiceAsFileTo_File_Click(sender As Object, e As EventArgs) Handles btnExistingTaskChoiceAsFileTo_File.Click
        'feladatválasztó ablakot hívja meg, a filing a meghívott ablak bezárása után nyílik meg a FileWithChosenTaskMetadata útján
        Dim TaskChooser As New SPTaxonomy_wWForms_TreeView.SPTreeview
        'AddHandler TaskChooser.ShowTaskSelected, AddressOf FileWithChosenTaskMetadata
        'TaskChooser.Show()
        TaskChooser.ShowDialog(Me)
        If Not TaskChooser.ShowDialog = DialogResult.OK Then Exit Sub
        FileWithChosenTaskMetadata(TaskChooser.SelectedTaskID)
    End Sub
    ''' <summary>
    ''' Nem használt btnExistingTaskChoiceAsFileTo_File végrehajtó része
    ''' </summary>
    ''' <param name="taskID"></param>
    Private Async Sub FileWithChosenTaskMetadata(taskID As Integer)
        Dim MySPHelper = Globals.ThisAddIn.Connection
        Dim SourceTask As TaskClass = Globals.ThisAddIn.Connection.Tasks.Where(Function(x) x.ID = taskID)
        Dim InterimDoc As DocClass = GetDocClassFromSelectedTask(SourceTask, ReturnSavedFileName())
        Dim FileLeafTarget As String = CheckUnicityFileNameInFolder(MySPHelper.Context, InterimDoc.PathtoSaveTo, InterimDoc.FileName)
        Dim _ListItem As CSOM.ListItem = Await CreateAndFileAsFile(FileLeafTarget, InterimDoc)
        MySPHelper.SetRelatedItemForTaskSecond(MySPHelper.Context, _ListItem, SourceTask.ID, FileLeafTarget)
        AddToHistoryFromSelectedTaskAndClear(SourceTask)
        'Nyissa meg egy böngészőablakban az új file-t!
        DocClass.OpenDocToEdit(_ListItem.Id)
    End Sub
#End Region

#Region "AfterSent és OnItemSend által meghívott végrehajtó részek, a formregion-ön szereplő adatok alapján"
    ''' <summary>
    ''' Válasz lefűzésével kapcsolatos, az AfterSent és OnItemSend által meghívott, az SP-re mentéssel kapcsolatos végrehajtó rész
    ''' </summary>
    ''' <param name="InterimDoc"></param>
    ''' <param name="taskID"></param>
    Public Shared Async Sub FileToDocLibraryAsFile(InterimDoc As DocClass, Optional taskID As Integer = 0)
        Dim MySPHelper = Globals.ThisAddIn.Connection
        Dim FileLeafTarget As String = CheckUnicityFileNameInFolder(MySPHelper.Context, InterimDoc.PathtoSaveTo, InterimDoc.FilePathLocalForUploading)
        Dim DocLibrarytoUpload = DocClass.ClientDocumentLibraryName
        Dim _ListItem As CSOM.ListItem
        Using fstream As New FileStream(InterimDoc.FilePathLocalForUploading, FileMode.Open)
            _ListItem = Await SPHelper.SaveToDocLibrary.SaveAndUploadAsync(Globals.ThisAddIn.Connection.Context, fstream, FileLeafTarget, DocLibrarytoUpload)
        End Using
        MySPHelper.SetListItemValuestoDocumentClassValues(_ListItem, InterimDoc)
        If taskID <> 0 Then
            MySPHelper.SetRelatedItemForTaskSecond(MySPHelper.Context, _ListItem, taskID, FileLeafTarget)
        End If
    End Sub
    ''' <summary>
    ''' Válasz lefűzésével kapcsolatos, az AfterSent és OnItemSend által meghívott végrehajtó rész metaadatainak összegyűjtésével kapcsolatos rész
    ''' </summary>
    ''' <param name="_record"></param>
    ''' <param name="localPathName"></param>
    ''' <returns></returns>
    Public Shared Function FillDocClassWithRecord(_record As ThisAddIn.ControlValuesGiven, localPathName As String) As DocClass
        Dim NewDoc As New DocClass With
                {.FilePathLocalForUploading = localPathName, .FileName = Path.GetFileName(localPathName),
                .TitleOrTaskName = _record.Title, .PathtoSaveTo = _record.Path, .Matter = _record.Matter, .Persons = _record.Partners}
        If Not IsNothing(_record.Task) Then
            With NewDoc
                If Not IsNothing(_record.Task.Matter) Then .Matter = _record.Task.Matter
                If Not IsNothing(_record.Task.Persons) Then .Persons = _record.Task.Persons
                If Not IsNothing(_record.Task.TitleOrTaskName) Then .TitleOrTaskName = _record.Task.TitleOrTaskName
                If Not IsNothing(_record.Task.ProjectSystems) Then .ProjectSystems = _record.Task.ProjectSystems
                If Not IsNothing(_record.Task.Keywords) Then .Keywords = _record.Task.Keywords
            End With
        End If
        Return NewDoc
    End Function
#End Region
    ''' <summary>
    ''' btnOpenFolder_Click és btnCreateFolderIfNotExisting_Click által meghívott segédfüggvény
    ''' </summary>
    ''' <returns></returns>
    Private Function ReturnUrlToOpen() As String
        Dim GetTargetUrlFolder = SPHelper.SPFileFolder.FilePathValidator("/" & DocClass.ClientDocumentLibraryFileRef & "/" & tbPathToSaveTo.Text)
        'Dim TargetLibrary = DocClass.ClientDocumentLibraryName
        Dim URLToOpen = My.Settings.spUrl & GetTargetUrlFolder
        Return URLToOpen
    End Function
    ''' <summary>
    ''' A TaskDoc controlok értékei alapján lementi a megfelelő fájlt egy ideiglenes könyvtárba, és visszaadja a mentett fájl útvonalát és nevét
    ''' </summary>
    ''' <returns>Az ideiglenes könyvtárba lementett fájl útvonala és neve</returns>
    Private Function ReturnSavedFileName() As String
        Dim SavedFileName = String.Empty
        If IsNothing(cbFileEmailOrAttachments.SelectedItem) OrElse cbFileEmailOrAttachments.SelectedItem.ID = AttachmentIndex.FullEmail Then
            SavedFileName = GetMailwAttachmentsIncludedAsFile(CurrentMail)
        ElseIf cbFileEmailOrAttachments.SelectedItem.ID = AttachmentIndex.EmailWithoutAttachment Then
            SavedFileName = GetMailWOAttachmentsAsFile(CurrentMail)
        Else
            Dim AttachmentSelected As OutlookItemAttachment = cbFileEmailOrAttachments.SelectedItem
            SavedFileName = GetMailAttachmentWOMailAsFile(CurrentMail, AttachmentSelected.ID)
        End If
        Return SavedFileName
    End Function
    ''' <summary>
    ''' Egy helyi meghajtóra lementett fájlt feltölti a megadott SP útvonalra. Végrehajtó rész FileToDocLibrary és a CreateNewTask_Tab függvényekhez
    ''' </summary>
    ''' <param name="filePathAtSpToSaveTo">SP cél címe (fileleaf típus) </param>
    ''' <param name="docToSave">DocClass típusú módon a feltöltendő adat (benne helyi útvonal és metaadatok is)</param>
    ''' <returns></returns>
    Private Async Function CreateAndFileAsFile(filePathAtSpToSaveTo As String, docToSave As DocClass) As Threading.Tasks.Task(Of CSOM.ListItem)
        Dim MySPHelper = Globals.ThisAddIn.Connection
        Dim DocLibrarytoUpload = DocClass.ClientDocumentLibraryName

        Dim _ListItem As CSOM.ListItem
        Using fstream As New FileStream(docToSave.FilePathLocalForUploading, FileMode.Open)
            _ListItem = Await SPHelper.SaveToDocLibrary.SaveAndUploadAsync(Globals.ThisAddIn.Connection.Context, fstream, filePathAtSpToSaveTo, DocLibrarytoUpload)
        End Using
        MySPHelper.SetListItemValuestoDocumentClassValues(_ListItem, docToSave)
        Return _ListItem
    End Function
#Region "History comboboxokba lefűzi meglévő adatokat és törli a formregion tartalmát"
    Private Sub AddToHistoryFromFileAndClear()
        cbFileHistory.Items.Add(New HistoryItem With {.Matter = cbMatter.SelectedItem, .Partner = lbTotalPartners.Items.Cast(Of PersonClass).ToList})
        If cbFileHistory.Items.Count > 20 Then cbFileHistory.Items.RemoveAt(0)
        cbMatter.SelectedItem = Nothing
        lbTotalPartners.Items.Clear()
    End Sub
    Private Sub AddToHistoryFromSelectedTaskAndClear(sourceTask As TaskClass)
        cbFileHistory.Items.Add(New HistoryItem With {.Task = sourceTask})
        If cbFileHistory.Items.Count > 20 Then cbFileHistory.Items.RemoveAt(0)
        cbTaskChosenHistoryNewTask.Items.Add(New HistoryItem With {.Task = sourceTask})
        If cbTaskChosenHistoryNewTask.Items.Count > 20 Then cbTaskChosenHistoryNewTask.Items.RemoveAt(0)
        cbMatter.SelectedItem = Nothing
        lbTotalPartners.Items.Clear()
    End Sub
#End Region

    ''' <summary>
    ''' Megadott helyi meghajtós fájlnévvel kapcsolatosan rögzítse a metaadatokat a File tab-ben foglalt controlok adatai alapján
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    Private Function GetDocClassFromFileTab(fileName As String) As DocClass
        Dim InterimDoc As New DocClass With {.FilePathLocalForUploading = fileName, .FileName = Path.GetFileName(fileName),
            .TitleOrTaskName = tbTitleFile.Text, .PathtoSaveTo = tbPathToSaveTo.Text}
        If lbTotalPartners.Items.Count > 0 Then
            InterimDoc.Persons = lbTotalPartners.Items.Cast(Of PersonClass).ToList
        End If
        InterimDoc = GetDataFromHistoryItem(InterimDoc, cbFileHistory)
        Return InterimDoc
    End Function
    Public Function IsTaskSelectedAsFileHistoryItem() As Boolean
        Dim SelectedHistory As HistoryItem = cbFileHistory.SelectedItem
        If IsNothing(SelectedHistory) Then Return False
        If Not IsNothing(SelectedHistory.Task) Then Return True Else Return False
    End Function
    Public Function ReturnTaskFromHistoryItem(cbHistory As ComboBox) As TaskClass
        Dim SelectedHistory As HistoryItem = cbHistory.SelectedItem
        If IsNothing(SelectedHistory) Then Return Nothing
        Return SelectedHistory.Task
    End Function
    Private Function GetDataFromHistoryItem(doc As DocClass, cbHistory As ComboBox) As DocClass
        If IsNothing(cbHistory.SelectedItem) Then
            doc.Matter = cbMatter.SelectedItem
        Else
            Dim SelectedHistory As HistoryItem = cbHistory.SelectedItem
            If Not IsNothing(SelectedHistory.Matter) Then doc.Matter = SelectedHistory.Matter
            If Not IsNothing(SelectedHistory.Partner) Then doc.Persons = SelectedHistory.Partner
            If Not IsNothing(SelectedHistory.Task) Then
                Dim SelectedTask As TaskClass = SelectedHistory.Task
                doc.Matter = SelectedTask.Matter
                doc.Persons = SelectedTask.Persons
                doc.Keywords = SelectedTask.Keywords
                doc.ProjectSystems = SelectedTask.ProjectSystems
            End If
        End If
        Return doc
    End Function
    Private Function GetDocClassFromSelectedTask(sourceTask As TaskClass, fileName As String) As DocClass
        Dim InterimDoc As New DocClass With {.FilePathLocalForUploading = fileName, .FileName = Path.GetFileName(fileName),
            .TitleOrTaskName = tbTitleFile.Text, .PathtoSaveTo = tbPathToSaveTo.Text}
        If lbTotalPartners.Items.Count > 0 Then
            InterimDoc.Persons = lbTotalPartners.Items.Cast(Of PersonClass).ToList
        End If
        InterimDoc = GetDataFromSelectedTask(InterimDoc, sourceTask)
        Return InterimDoc
    End Function
    Private Function GetDataFromSelectedTask(doc As DocClass, SourceTask As TaskClass) As DocClass
        If Not IsNothing(cbMatter.SelectedItem) Then doc.Matter = cbMatter.SelectedItem
        If IsNothing(SourceTask) Then Return doc
        If Not IsNothing(SourceTask.Matter) Then doc.Matter = SourceTask.Matter
        If Not IsNothing(SourceTask.Persons) Then doc.Persons = SourceTask.Persons
        doc.Keywords = SourceTask.Keywords
        doc.ProjectSystems = SourceTask.ProjectSystems
        Return doc
    End Function
    Private Function GetDocClassFromTaskHistory(fileName As String) As DocClass
        Dim InterimDoc As New DocClass With {.FilePathLocalForUploading = fileName, .FileName = Path.GetFileName(fileName),
            .TitleOrTaskName = tbTitleFile.Text, .PathtoSaveTo = tbPathToSaveTo.Text}
        If lbTotalPartners.Items.Count > 0 Then
            InterimDoc.Persons = lbTotalPartners.Items.Cast(Of PersonClass).ToList
        End If
        InterimDoc = GetDataFromHistoryItem(InterimDoc, cbTaskChosenHistoryNewTask)
        Return InterimDoc
    End Function
#Region "Create and set new tasks: using MailItem, CSOM and TaskClass"
    Private Function CreateNewTask() As TaskClass
        Dim MySPHelper = Globals.ThisAddIn.Connection
        Dim NewTaskAsListItem = SetTaskListItemFromMail(CurrentMail, context:=MySPHelper.Context)
        NewTaskAsListItem.Update()
        MySPHelper.Context.ExecuteQuery()
        AddCategoryToMail(CurrentMail, TaskClass.TaskedToSP)
        TaskClass.OpenTaskToEdit(NewTaskAsListItem.Id)
        Globals.ThisAddIn.LoadAndUpdateTasks()
        Dim NewTaskClass = Globals.ThisAddIn.Connection.Tasks.Where(Function(x) x.ID = NewTaskAsListItem.Id)
        Return NewTaskClass
    End Function
    Private Sub CreateNewTaskBasedOnOld(taskToBaseOn As TaskClass)
        If IsNothing(taskToBaseOn) Then
            MsgBox("Nem létező task alapján kellene létrehozni")
            Exit Sub
        End If

        Dim MySPHelper = Globals.ThisAddIn.Connection
        Dim NewTaskAsListItem = SetTaskListItemFromMail(CurrentMail, context:=MySPHelper.Context)
        'CurrentMail itemek adatainak felülírása Task szerinti adatokkal
        TaskClass.SetTaskListItemFromOldTask(MySPHelper, tbTitleFile.Text, NewTaskAsListItem, taskToBaseOn)
        Dim ParentID As Integer = 0
        If cbSubTask.Checked Then ParentID = taskToBaseOn.ID
        SetParentIDAndLoad(NewTaskAsListItem, ParentID)
        AddCategoryToMail(CurrentMail, TaskClass.TaskedToSP)
        TaskClass.OpenTaskToEdit(NewTaskAsListItem.Id)
        Globals.ThisAddIn.LoadAndUpdateTasks()
    End Sub

    Public Function SetTaskListItemFromMail(_MailItem As MailItem, Optional targetListItem As CSOM.ListItem = Nothing, Optional context As CSOM.ClientContext = Nothing) As CSOM.ListItem
        If IsNothing(context) Then
            context = Globals.ThisAddIn.Connection.Context
        End If
        If IsNothing(targetListItem) Then
            targetListItem = TaskClass.NewTaskListItem(context)
        End If

        targetListItem("EmConversationID") = _MailItem.ConversationID
        targetListItem("EmID") = _MailItem.EntryID
        Select Case _MailItem.BodyFormat
            Case OlBodyFormat.olFormatHTML, OlBodyFormat.olFormatUnspecified
                targetListItem("EmBody") = _MailItem.HTMLBody
            Case OlBodyFormat.olFormatPlain
                targetListItem("EmBody") = _MailItem.Body
            Case OlBodyFormat.olFormatRichText
                targetListItem("EmBody") = _MailItem.RTFBody
        End Select
        targetListItem("EmType") = _MailItem.BodyFormat
        targetListItem("EmCC") = _MailItem.CC
        targetListItem("EmCCSMTPAddress") = String.Join("; ", GetAllAddress(_MailItem, Microsoft.Office.Interop.Outlook.OlMailRecipientType.olCC))
        targetListItem("EmFrom") = _MailItem.SenderEmailAddress
        targetListItem("EmFromName") = _MailItem.SenderName
        targetListItem("EmSubject") = _MailItem.Subject
        targetListItem("EmTo") = _MailItem.To
        targetListItem("EmToAddress") = String.Join("; ", GetAllAddress(_MailItem, Microsoft.Office.Interop.Outlook.OlMailRecipientType.olTo))
        targetListItem("EmSensitivity") = _MailItem.Sensitivity
        targetListItem("Priority") = _MailItem.Priority
        Return targetListItem
    End Function
    Private Sub SetParentIDAndLoad(ByRef targetListItem As CSOM.ListItem, ParentID As Integer)
        Dim MySPHelper = Globals.ThisAddIn.Connection
        If ParentID > 0 Then
            targetListItem("ParentID") = MySPHelper.GetLookupValue(ParentID)
        End If
        MySPHelper.Context.Load(targetListItem)
        targetListItem.Update()
        MySPHelper.Context.ExecuteQuery()
    End Sub
#End Region
    'Private Function GetTaxonomyFields(Optional ByRef list As CSOM.List = Nothing) As CSOM.Taxonomy.TaxonomyField
    '    Dim MySPHelper = Globals.ThisAddIn.Connection
    '    list = MySPHelper.Context.Web.Lists.GetByTitle(DocClass.ClientDocumentLibraryName)
    '    MySPHelper.Context.Load(list)
    '    MySPHelper.Context.Load(list, Function(dl) dl.ContentTypes, Function(dl) dl.ParentWeb.Id, Function(dl) dl.Id)
    '    MySPHelper.Context.Load(list.RootFolder.Files)
    '    MySPHelper.Context.Load(list.Fields)
    '    Dim taxFieldEKW As CSOM.Taxonomy.TaxonomyField = MySPHelper.Context.CastTo(Of Microsoft.SharePoint.Client.Taxonomy.TaxonomyField)(list.Fields.GetByInternalNameOrTitle("TaxKeyword"))
    '    MySPHelper.Context.Load(taxFieldEKW)
    '    MySPHelper.Context.ExecuteQuery()
    '    Return taxFieldEKW
    'End Function
End Class
