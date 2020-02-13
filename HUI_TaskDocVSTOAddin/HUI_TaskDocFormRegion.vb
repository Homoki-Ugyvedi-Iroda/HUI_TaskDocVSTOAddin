Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports SPHelper.SPHUI
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

#End Region
    Property LastTabPage As Integer
    Property ShowAllPartners As Boolean
    Property TaskTreeView As TreeView
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
    'Occurs before the form region is displayed. 
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub HUI_TaskDocFormRegion_FormRegionShowing(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionShowing
        'Feltölti a három cb-t értékekkel, mindhárom cb értéke azonos, csak a Taskból és a body-ból veszi ki a history-t
        CurrentMail = TryCast(Me.OutlookItem, MailItem)
        If IsNothing(CurrentMail) Then Exit Sub
        If OutlookApplicationHelper.IsNewAndEmpty(CurrentMail) Then Exit Sub
        tbTitleFile.Text = CurrentMail.Subject
        LoadAttachmentNames(CurrentMail)
        LoadValuesIntocbTaskChosen({cbTaskChosenHistoryNewTask, cbTaskChosenHistoryFileTask, cbFileHistory}, CurrentMail)
        LoadValuesIncbMattersAllActive()
        LoadValuesIncbPartnersAllActive()
        LoadValuesInTaskChooser()
    End Sub


    'Occurs when the form region is closed.   
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub HUI_TaskDocFormRegion_FormRegionClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionClosed

    End Sub
    Private Sub LoadValuesInTaskChooser()
        Dim MySPTaxonomyTreeView As New SPTaxonomy_wWForms_TreeView.SPTreeview
        Me.TaskTreeView = MySPTaxonomyTreeView.ShowNodeswithParents(Globals.ThisAddIn.Connection.Tasks, True)

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
        cbPartner.DisplayMember = "Title"
        cbPartner.ValueMember = "ID"
        lbTotalPartners.DisplayMember = "Title"
        lbTotalPartners.ValueMember = "ID"
        cbMatter.Items.AddRange(Globals.ThisAddIn.Connection.Persons.Where(Function(x) x.Active = True).ToArray)
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

#Region "CreateNewTask"

    Private Async Sub btnHistoryChosenAsTemplateForNewTask_Click(sender As Object, e As EventArgs) Handles btnHistoryChosenAsTemplateForNewTask.Click
        If IsNothing(cbTaskChosenHistoryNewTask.SelectedValue) Then Exit Sub
        Dim SavedFileName As String = ReturnSavedFileName()
        Dim MySPHelper = Globals.ThisAddIn.Connection
        Dim DocLibrarytoUpload As CSOM.List = MySPHelper.Context.Web.Lists.GetByTitle(DocClass.ClientDocumentLibraryName)
        MySPHelper.Context.Load(DocLibrarytoUpload)
        MySPHelper.Context.Load(DocLibrarytoUpload, Function(dl) dl.ContentTypes, Function(dl) dl.ParentWeb.Id, Function(dl) dl.Id)
        MySPHelper.Context.Load(DocLibrarytoUpload.RootFolder.Files)
        MySPHelper.Context.Load(DocLibrarytoUpload.Fields)
        Dim taxFieldEKW = MySPHelper.Context.CastTo(Of Microsoft.SharePoint.Client.Taxonomy.TaxonomyField)(DocLibrarytoUpload.Fields.GetByInternalNameOrTitle("TaxKeyword"))
        MySPHelper.Context.Load(taxFieldEKW)
        MySPHelper.Context.ExecuteQuery()

        Dim InterimDoc As DocClass = GetDocClassFromFileTab(SavedFileName)
        Dim FileLeafTarget As String
        Dim _ListItem As Microsoft.SharePoint.Client.ListItem
        Using fstream As New FileStream(InterimDoc.FilePathLocalForUploading, FileMode.Open)
            FileLeafTarget = CheckUnicity(InterimDoc.PathtoSaveTo, InterimDoc.FileName)
            _ListItem = Await SPHelper.SaveToDocLibrary.SaveAndUploadAsync(Globals.ThisAddIn.Connection.Context, fstream, FileLeafTarget, DocLibrarytoUpload)
            MySPHelper.SetListItemValuestoDocumentClassValues(_ListItem, InterimDoc, taxFieldEKW, taxFieldInternalDoc:=Nothing)
        End Using
        If IsTaskSelectedAsFileHistoryItem() Then
            '#Ha volt korábbi task kiválasztva a cbFileHistory-ban, akkor fűzze hozzá a taskhoz related item-ként az új itemet
            MySPHelper.SetRelatedItemForTaskSecond(MySPHelper.Context, _ListItem, ReturnTaskFromHistoryItem(cbTaskChosenHistoryNewTask).ID, FileLeafTarget)
            'MySPHelper.SetRelatedItemForTask(MySPHelper.Context, ReturnTaskFromFilehistoryItem.ID, FileLeafTarget)
        End If
        '###Mentse le egy új taskként a korábbi adataival, majd nyissa meg egy böngészőablakban az új taskot!
    End Sub

    Private Sub btnExistingTaskChoiceAsTemplateForNewTask_Click(sender As Object, e As EventArgs) Handles btnExistingTaskChoiceAsTemplateForNewTask.Click
        '#feladatválasztó ablakot hívja meg, újra gyártsa le a faszerkezetet, ha újat hozunk létre
    End Sub

    Private Sub btnCreateNewTask_Click(sender As Object, e As EventArgs) Handles btnCreateNewTask.Click
        '# létrehoz egy új üres taskot, amihez csatolja ezt az emailt/csatolmányt és annak megfelelő, SP szerinti új task ablak megnyitása? 
        '#feladatválasztó ablakot hívja meg, újra gyártsa le a faszerkezetet, ha újat hozunk létre
    End Sub
#End Region
#Region "FileToDocLibrary"
    Private Function ReturnUrlToOpen() As String
        Dim GetTargetUrlFolder = "/" & DocClass.ClientDocumentLibraryFileRef & "/" & tbPathToSaveTo.Text
        Dim TargetLibrary = DocClass.ClientDocumentLibraryName
        Dim URLToOpen = SPHelper.SPFileFolder.FilePathValidator(My.Settings.spUrl & GetTargetUrlFolder)
        Return URLToOpen
    End Function
    Private Sub btnAddPartnerFile_Click(sender As Object, e As EventArgs) Handles btnAddPartnerFile.Click
        If IsNothing(cbPartner.SelectedValue) Then Exit Sub
        lbTotalPartners.Items.Add(cbPartner.SelectedItem)
        CheckIfPathChangesForPartnerChange(sender, e)
    End Sub
    Private Sub btnDeleteLastPartner_Click(sender As Object, e As EventArgs) Handles btnDeleteLastPartner.Click
        If lbTotalPartners.Items.Count < 1 Then Exit Sub
        lbTotalPartners.Items.Remove(lbTotalPartners.Items.Count - 1)
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
        Dim DefaultPartner As PersonClass = Globals.ThisAddIn.Connection.Persons.Where(Function(x) x.Id = SelectedMatter.MatterDefaultPerson)
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
    Private Sub CheckIfPathChangesForPartnerChange(sender As Object, e As EventArgs)
        cbMatter_SelectedIndexChanged(sender, e)
    End Sub
    Private Sub btnChangePartnerOrder_Click(sender As Object, e As EventArgs) Handles btnChangePartnerOrder.Click
        If cbPartner.Items.Count < 2 Then Exit Sub
        Dim InterimItem As Object = cbPartner.Items(0)
        cbPartner.Items.Remove(0)
        cbPartner.Items.Add(InterimItem)
        CheckIfPathChangesForPartnerChange(sender, e)
    End Sub

    Private Async Sub btnFileToDocLibrary_Click(sender As Object, e As EventArgs) Handles btnFileToDocLibrary.Click
        Dim SavedFileName As String = ReturnSavedFileName()
        Dim MySPHelper = Globals.ThisAddIn.Connection
        Dim DocLibrarytoUpload As CSOM.List = MySPHelper.Context.Web.Lists.GetByTitle(DocClass.ClientDocumentLibraryName)
        MySPHelper.Context.Load(DocLibrarytoUpload)
        MySPHelper.Context.Load(DocLibrarytoUpload, Function(dl) dl.ContentTypes, Function(dl) dl.ParentWeb.Id, Function(dl) dl.Id)
        MySPHelper.Context.Load(DocLibrarytoUpload.RootFolder.Files)
        MySPHelper.Context.Load(DocLibrarytoUpload.Fields)
        Dim taxFieldEKW = MySPHelper.Context.CastTo(Of Microsoft.SharePoint.Client.Taxonomy.TaxonomyField)(DocLibrarytoUpload.Fields.GetByInternalNameOrTitle("TaxKeyword"))
        'nem kell, mivel InternalDoc nincsen mentve itt (tasknál nincsen):
        '   Dim taxFieldInternalDoc = MySPHelper.Context.CastTo(Of Microsoft.SharePoint.Client.Taxonomy.TaxonomyField)(DocLibrarytoUpload.Fields.GetByInternalNameOrTitle(InternalDocumentTypeColumnInternalName))
        '   MySPHelper.Context.Load(taxFieldInternalDoc)
        MySPHelper.Context.Load(taxFieldEKW)
        MySPHelper.Context.ExecuteQuery()

        Dim InterimDoc As DocClass = GetDocClassFromFileTab(SavedFileName)
        Dim FileLeafTarget As String
        Dim _ListItem As Microsoft.SharePoint.Client.ListItem
        Using fstream As New FileStream(InterimDoc.FilePathLocalForUploading, FileMode.Open)
            FileLeafTarget = CheckUnicity(InterimDoc.PathtoSaveTo, InterimDoc.FileName)
            _ListItem = Await SPHelper.SaveToDocLibrary.SaveAndUploadAsync(Globals.ThisAddIn.Connection.Context, fstream, FileLeafTarget, DocLibrarytoUpload)
            MySPHelper.SetListItemValuestoDocumentClassValues(_ListItem, InterimDoc, taxFieldEKW, taxFieldInternalDoc:=Nothing) ', KeywordsFile:=InterimDoc.Keywords)
        End Using
        If IsTaskSelectedAsFileHistoryItem() Then
            '#Ha volt korábbi task kiválasztva a cbFileHistory-ban, akkor fűzze hozzá a taskhoz related item-ként az új itemet
            MySPHelper.SetRelatedItemForTaskSecond(MySPHelper.Context, _ListItem, ReturnTaskFromHistoryItem(cbFileHistory).ID, FileLeafTarget)
            'MySPHelper.SetRelatedItemForTask(MySPHelper.Context, ReturnTaskFromFilehistoryItem.ID, FileLeafTarget)
        End If
        AddToHistoryFromFileAndClear()
    End Sub

    Private Sub btnTaskChoiceAsFileTo_File_Click(sender As Object, e As EventArgs) Handles btnExistingTaskChoiceAsFileTo_File.Click
        '#feladatválasztó ablakot hívja meg
        Dim TaskChooser As New SPTaxonomy_wWForms_TreeView.SPTreeview
        AddHandler TaskChooser.ShowTaskSelected, AddressOf FileWithChosenTaskMetadata
        TaskChooser.Show()
    End Sub

    Private Async Sub FileWithChosenTaskMetadata(taskID As Integer)
        Dim SavedFileName As String = ReturnSavedFileName()
        Dim MySPHelper = Globals.ThisAddIn.Connection

        Dim DocLibrarytoUpload As CSOM.List = MySPHelper.Context.Web.Lists.GetByTitle(DocClass.ClientDocumentLibraryName)
        MySPHelper.Context.Load(DocLibrarytoUpload)
        MySPHelper.Context.Load(DocLibrarytoUpload, Function(dl) dl.ContentTypes, Function(dl) dl.ParentWeb.Id, Function(dl) dl.Id)
        MySPHelper.Context.Load(DocLibrarytoUpload.RootFolder.Files)
        MySPHelper.Context.Load(DocLibrarytoUpload.Fields)
        Dim taxFieldEKW = MySPHelper.Context.CastTo(Of Microsoft.SharePoint.Client.Taxonomy.TaxonomyField)(DocLibrarytoUpload.Fields.GetByInternalNameOrTitle("TaxKeyword"))
        MySPHelper.Context.Load(taxFieldEKW)
        'kellenek ezek?
        MySPHelper.Context.ExecuteQuery()
        Dim SourceTask As TaskClass = Globals.ThisAddIn.Connection.Tasks.Where(Function(x) x.ID = taskID)
        Dim InterimDoc As DocClass = GetDocClassFromSelectedTask(SourceTask, SavedFileName)
        Dim FileLeafTarget As String
        Dim _ListItem As Microsoft.SharePoint.Client.ListItem
        Using fstream As New FileStream(InterimDoc.FilePathLocalForUploading, FileMode.Open)
            FileLeafTarget = CheckUnicity(InterimDoc.PathtoSaveTo, InterimDoc.FileName)
            _ListItem = Await SPHelper.SaveToDocLibrary.SaveAndUploadAsync(Globals.ThisAddIn.Connection.Context, fstream, FileLeafTarget, DocLibrarytoUpload)
            MySPHelper.SetListItemValuestoDocumentClassValues(_ListItem, InterimDoc, taxFieldEKW, taxFieldInternalDoc:=Nothing)
        End Using
        MySPHelper.SetRelatedItemForTaskSecond(MySPHelper.Context, _ListItem, SourceTask.ID, FileLeafTarget)
        AddToHistoryFromSelectedTaskAndClear(SourceTask)
    End Sub

    Private Function CheckUnicity(targetPath As String, targetFileName As String) As String
        Dim ReturnName As String = targetPath & targetFileName
        If SPHelper.SPFileFolder.TryGetFolderByServerRelativeUrl(Globals.ThisAddIn.Connection.Context, targetPath & targetFileName) Then
            Dim random = HPHelper.StringRelated.RandomCharGet(5)
            Dim msgResponse As MsgBoxResult = MsgBox("The filename " & targetFileName & " already exists in the folder." & Environment.NewLine &
                                                     "Do you want to overwrite (YES) or " & Environment.NewLine & "save the name with a new ending (NO = " & random & ")?", MsgBoxStyle.YesNo)
            If msgResponse = vbNo Then ReturnName = targetPath & Path.GetFileNameWithoutExtension(targetFileName) & random & Path.GetExtension(targetFileName)
        End If
        Return ReturnName
    End Function
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

    Private Sub cbFileEmailOrAttachments_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbFileEmailOrAttachments.SelectedValueChanged
        If cbFileEmailOrAttachments.SelectedIndex < 2 Then Exit Sub
        Dim SelectedAttachment As Attachment = cbFileEmailOrAttachments.SelectedItem
        tbTitleFile.Text = SelectedAttachment.DisplayName
    End Sub

#End Region
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

    Private Function GetDocClassFromFileTab(fileName As String) As DocClass
        Dim InterimDoc As New DocClass With {.FilePathLocalForUploading = fileName, .FileName = Path.GetFileName(fileName),
            .TitleOrTaskName = tbTitleFile.Text, .PathtoSaveTo = tbPathToSaveTo.Text}
        If lbTotalPartners.Items.Count > 0 Then
            InterimDoc.Persons = lbTotalPartners.Items.Cast(Of PersonClass).ToList
        End If
        InterimDoc = GetDataFromHistoryItem(InterimDoc, cbFileHistory)
        Return InterimDoc
    End Function
    Private Function IsTaskSelectedAsFileHistoryItem() As Boolean
        Dim SelectedHistory As HistoryItem = cbFileHistory.SelectedItem
        If IsNothing(SelectedHistory) Then Return False
        If Not IsNothing(SelectedHistory.Task) Then Return True Else Return False
    End Function
    Private Function ReturnTaskFromHistoryItem(cbHistory As ComboBox) As TaskClass
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

    Private Sub cbTaskChosenHistoryNewTask_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbTaskChosenHistoryNewTask.SelectedValueChanged
        btnHistoryChosenAsTemplateForNewTask.Enabled = True
    End Sub
End Class
