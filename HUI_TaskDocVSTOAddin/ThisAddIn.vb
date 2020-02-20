Imports System.IO
Imports SPHelper.SPHUI
Imports System.Windows.Forms
Imports System.Threading.Tasks

Public Class ThisAddIn
    Property Connection As SPHelper.SPHUI
    Property TempPath As String

    Public Const UnknownFolderName = "/_Unknown/"
    Public Const DeferredDeliveryTimeinSeconds As Integer = 30

    Public Log As HPHelper.DebugTesting
    Public WithEvents TimerForRefresh3min As New Timer
    Private WithEvents SentItem As Outlook.Items

    Public CacheDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Cached")
    Public CInternalDocTypesPath = Path.Combine(CacheDirectory, "InternalDocTypes.json")
    Public CNonWorkDocTypesPath = Path.Combine(CacheDirectory, "NonWorkDocTypes.json")
    Public CWorkDocTypesPath = Path.Combine(CacheDirectory, "WorkDocTypes.json")
    Public CMattersPath = Path.Combine(CacheDirectory, "Matters.json")
    Public CPersonsPath = Path.Combine(CacheDirectory, "Persons.json")
    Public CPrioritiesPath = Path.Combine(CacheDirectory, "Priorities.json")
    Public CTasksPath = Path.Combine(CacheDirectory, "Tasks.json")
    Public CTermsPath = Path.Combine(CacheDirectory, "Terms.json")
    Public CTaskTypesPath = Path.Combine(CacheDirectory, "TaskTypes.json")
    Public CUsersPath = Path.Combine(CacheDirectory, "Users.json")
    Public CKeywords = Path.Combine(CacheDirectory, "Keywords.json")
    Public CTaxFieldnames = Path.Combine(CacheDirectory, "TaxFields.json")
    Public RefreshCounter As Byte = 1
    Private LastItemFiledGuid As String

    Public Class ControlValuesGiven
        Property Task As TaskClass
        Property Matter As MatterClass
        Property Title As String
        Property Path As String
        Property Partners As List(Of PersonClass)
        'Property CreateSubTask As Boolean
        'Property CreateTask As Boolean
    End Class

    Private Function GetCachePath(name As String) As String
        Return Path.Combine(CacheDirectory, name + ".json")
    End Function

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Logstart()
        SetAndVerifyTempPath()
        StartSpLoadCacheSpRefresh()
        TimerForRefresh3min.Interval = 3 * 60 * 1000
        TimerForRefresh3min.Start()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Try
            Directory.Delete(TempPath, True)
        Catch ex As Exception
            Log.logger.Error(ex.Message & "Could not delete directory " & TempPath)
        End Try
    End Sub
    Private Sub Logstart()
        Dim logfile As String = "debugtrace_" & My.Application.Info.AssemblyName & ".log.txt"
        Dim logfilepath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), logfile)
        Log = New HPHelper.DebugTesting(logfilepath)
        Log.logger.Info(My.Application.Info.AssemblyName & " started")
    End Sub
    Private Async Sub StartSpLoadCacheSpRefresh()
        Me.Connection = New SPHelper.SPHUI()
        LoadCachedValues()
        Await RefreshSpLookupValues(False)
    End Sub
    Private Sub SetAndVerifyTempPath()
        Globals.ThisAddIn.TempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName)
        If String.IsNullOrWhiteSpace(Me.TempPath) Then
            Globals.ThisAddIn.TempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName)
        End If
        If Not Directory.Exists(Me.TempPath) Then Directory.CreateDirectory(Me.TempPath)
        Globals.ThisAddIn.Log.logger.Info("SetAndVerifyTempPath finished")
    End Sub

    Private Async Sub TimerForRefresh3min_Tick(sender As Object, e As EventArgs) Handles TimerForRefresh3min.Tick
        If RefreshCounter Mod 10 = 0 Then
            Await RefreshSpLookupValues(PartialRefresh:=False)
            SerializeValues(False)
        Else
            Await RefreshSpLookupValues(PartialRefresh:=True)
            SerializeValues(True)
        End If
        RefreshCounter += 1
    End Sub

    Public Async Function RefreshSpLookupValues(PartialRefresh As Boolean) As Task
        Dim m As New DataLayer()
        If Me.Connection.Connected = False Then
            Me.Connection.Connect(My.Settings.spUrl)
            If Me.Connection.Connected = False Then
                Globals.ThisAddIn.Log.logger.Info("No connection, no refresh")
                Exit Function
            End If
        End If
        If PartialRefresh = False Then
            Me.Connection.Users = Await m.GetAllUsers(Me.Connection)
            Me.Connection.Matters = Await m.GetAllMattersAsync(Me.Connection)
            Me.Connection.InternalDocTypes = m.GetAllInternalDocTypes(Me.Connection)
            Me.Connection.WorkDocumentType = Await m.GetAllWorkDocTypesAsync(Me.Connection)
            Await LoadChoiceColumns(m)
        End If
        Me.Connection.Persons = Await m.GetAllPersonsAsync(Me.Connection)
        Me.Connection.Tasks = Await m.GetAllTasksAsync(Me.Connection)
        Me.Connection.AllTerms = Await m.GetAllTermsAsync(Me.Connection)
        Me.Connection.Keywords = Await m.GetAllKeywords(Me.Connection)
        Me.Connection.taxFieldEKW = Await TermClass.GetTaxFieldEKWAsync(Connection.Context, TaskClass.ListName)
        Me.Connection.taxFieldInternalDoc = Await TermClass.GetFieldInternalDocAsync(Connection.Context)
        Globals.ThisAddIn.Log.logger.Info("RefreshSpLookupValues finished")
    End Function
    Private Async Function LoadChoiceColumns(input As DataLayer) As Task
        Dim DictionaryToLoad As Dictionary(Of String, Microsoft.SharePoint.Client.FieldMultiChoice) = Await input.GetAllColumnChoiceTypesAsync(Me.Connection)
        For Each _column In DictionaryToLoad
            If Not IsNothing(_column.Value) Then
                Select Case _column.Key
                    Case "Priority"
                        Me.Connection.Priorities = _column.Value.Choices.ToList
                    Case TaskClass.TaskTypeColumnInternalName
                        Me.Connection.TaskTypes = {String.Empty}.ToList
                        Me.Connection.TaskTypes.AddRange(_column.Value.Choices)
                    Case NonWorkDocTypeColumnInternalName_AsSiteColumn
                        Me.Connection.NonWorkDocTypes = {String.Empty}.ToList
                        Me.Connection.NonWorkDocTypes.AddRange(_column.Value.Choices)
                End Select
            End If
        Next
    End Function
    Private Sub LoadCachedValues()
        Me.Connection.AllTerms = DeserializeTerms(CTermsPath)
        Me.Connection.InternalDocTypes = DeserializeTerms(CInternalDocTypesPath)
        Me.Connection.Keywords = DeserializeTerms(CKeywords)
        Me.Connection.NonWorkDocTypes = DeserializeListofString(CNonWorkDocTypesPath)
        Me.Connection.TaskTypes = DeserializeListofString(CTaskTypesPath)
        Me.Connection.Priorities = DeserializeListofString(CPrioritiesPath)
        Me.Connection.Users = DeserializeUsers(CUsersPath)
        Me.Connection.WorkDocumentType = DeserializeWorkDocTypes(CWorkDocTypesPath)
        Me.Connection.Matters = DeserializeMatterClass(CMattersPath)
        Me.Connection.Persons = DeserializePersonClass(CPersonsPath)
        Me.Connection.Tasks = DeserializeTaskClass(CTasksPath)
        Dim listTaxField As List(Of Microsoft.SharePoint.Client.Taxonomy.TaxonomyField) = DeserializeFieldNames(CTaxFieldnames)
        If listTaxField.Count = 2 Then
            Me.Connection.taxFieldEKW = listTaxField.First
            Me.Connection.taxFieldInternalDoc = listTaxField.Last
        End If
        Globals.ThisAddIn.Log.logger.Info("Values loaded from cache.")
    End Sub
    Private Sub SerializeValues(PartialRefresh As Boolean)
        If Not Directory.Exists(CacheDirectory) Then
            Directory.CreateDirectory(CacheDirectory)
            Globals.ThisAddIn.Log.logger.Info("Cache directory created at " & CacheDirectory)
        End If
        Dim success As Byte = 0
        'Dim ObjectsToCheck = {Me.Connection.Keywords, Me.Connection.Users, Me.Connection.WorkDocumentType, Me.Connection.NonWorkDocTypes, Me.Connection.TaskTypes, Me.Connection.Priorities, Me.Connection.AllTerms}
        If PartialRefresh = False Then
            If Not IsNothing(Me.Connection.Users) AndAlso Me.Connection.Users.Count > 0 Then
                Dim succeeded = HPHelper.Serialize.Serialize(CUsersPath, Me.Connection.Users)
                If succeeded Then success += 1
            End If
            If Not IsNothing(Me.Connection.WorkDocumentType) AndAlso Me.Connection.WorkDocumentType.Count > 0 Then
                Dim succeeded = HPHelper.Serialize.Serialize(CWorkDocTypesPath, Me.Connection.WorkDocumentType)
                If succeeded Then success += 1
            End If
            If Not IsNothing(Me.Connection.NonWorkDocTypes) AndAlso Me.Connection.NonWorkDocTypes.Count > 0 Then
                Dim succeeded = HPHelper.Serialize.Serialize(CNonWorkDocTypesPath, Me.Connection.NonWorkDocTypes)
                If succeeded Then success += 1
            End If
            If Not IsNothing(Me.Connection.TaskTypes) AndAlso Me.Connection.TaskTypes.Count > 0 Then
                Dim succeeded = HPHelper.Serialize.Serialize(CTaskTypesPath, Me.Connection.TaskTypes)
                If succeeded Then success += 1
            End If
            If Not IsNothing(Me.Connection.Priorities) AndAlso Me.Connection.Priorities.Count > 0 Then
                Dim succeeded = HPHelper.Serialize.Serialize(CPrioritiesPath, Me.Connection.Priorities)
                If succeeded Then success += 1
            End If
            If Not IsNothing(Me.Connection.InternalDocTypes) AndAlso Me.Connection.InternalDocTypes.Count > 0 Then
                Dim succeeded = HPHelper.Serialize.Serialize(CInternalDocTypesPath, Me.Connection.InternalDocTypes)
                If succeeded Then success += 1
            End If
            If Not IsNothing(Me.Connection.Matters) AndAlso Me.Connection.Matters.Count > 0 Then
                Dim succeeded = HPHelper.Serialize.Serialize(CMattersPath, Me.Connection.Matters)
                If succeeded Then success += 1
            End If
        End If
        If Not IsNothing(Me.Connection.Keywords) AndAlso Me.Connection.Keywords.Count > 0 Then
            Dim succeeded = HPHelper.Serialize.Serialize(CKeywords, Me.Connection.Keywords)
            If succeeded Then success += 1
        End If
        If Not IsNothing(Me.Connection.AllTerms) AndAlso Me.Connection.AllTerms.Count > 0 Then
            Dim succeeded = HPHelper.Serialize.Serialize(CTermsPath, Me.Connection.AllTerms)
            If succeeded Then success += 1
        End If
        If Not IsNothing(Me.Connection.Tasks) AndAlso Me.Connection.Tasks.Count > 0 Then
            Dim succeeded = HPHelper.Serialize.Serialize(CTasksPath, Me.Connection.Tasks)
            If succeeded Then success += 1
        End If
        If Not IsNothing(Me.Connection.Persons) AndAlso Me.Connection.Persons.Count > 0 Then
            Dim succeeded = HPHelper.Serialize.Serialize(CPersonsPath, Me.Connection.Persons)
            If succeeded Then success += 1
        End If
        If Not IsNothing(Me.Connection.taxFieldEKW) AndAlso IsNothing(Me.Connection.taxFieldInternalDoc) Then
            Dim taxList As List(Of Microsoft.SharePoint.Client.Taxonomy.TaxonomyField) = New List(Of Microsoft.SharePoint.Client.Taxonomy.TaxonomyField)
            taxList.AddRange({Me.Connection.taxFieldEKW, Me.Connection.taxFieldInternalDoc})
            Dim succeeded = HPHelper.Serialize.Serialize(CTaxFieldnames, taxList)
        End If
        Globals.ThisAddIn.Log.logger.Info("Items serialized from memory to file (" & success & ")")
    End Sub

#Region "ItemSend"
    Private Function RecordFileSaveAsInformation(InputRegion As HUI_TaskDocFormRegion) As ControlValuesGiven
        Dim Result As New ControlValuesGiven
        If InputRegion.cbMatter.Items.Count > 0 Then
            If IsNothing(InputRegion.cbMatter.SelectedItem) Then
                Result.Matter = InputRegion.cbMatter.Items(0)
            Else Result.Matter = InputRegion.cbMatter.SelectedItem
            End If
        End If
        For Each item As PersonClass In InputRegion.lbTotalPartners.Items
            Result.Partners.Add(item)
        Next
        Result.Path = InputRegion.tbPathToSaveTo.Text
        Result.Title = InputRegion.tbTitleFile.Text
        With InputRegion.cbFileHistory
            If .Items.Count = 0 Or IsNothing(.SelectedItem) Then Return Result
            Dim SelectedHistory As HUI_TaskDocFormRegion.HistoryItem = InputRegion.cbFileHistory.SelectedItem
            If IsNothing(SelectedHistory) Then Return Result
            If Not IsNothing(SelectedHistory.Matter) Then Result.Matter = SelectedHistory.Matter
            If Not IsNothing(SelectedHistory.Partner) Then Result.Partners = SelectedHistory.Partner
            If Not IsNothing(SelectedHistory.Task) Then
                Result.Task = SelectedHistory.Task
            End If
        End With
        Return Result
    End Function
    Private Function RecordCreateNewTaskAsInformation(InputRegion As HUI_TaskDocFormRegion) As ControlValuesGiven
        Dim Result As New ControlValuesGiven
        With InputRegion.cbTaskChosenHistoryNewTask
            If IsNothing(.SelectedItem) Then
                Return RecordFileSaveAsInformation(InputRegion)
            Else
                Dim SelectedHistory As HUI_TaskDocFormRegion.HistoryItem = .SelectedItem
                If IsNothing(SelectedHistory) Then Return Result
                If Not IsNothing(SelectedHistory.Matter) Then Result.Matter = SelectedHistory.Matter
                If Not IsNothing(SelectedHistory.Partner) Then Result.Partners = SelectedHistory.Partner
                If Not IsNothing(SelectedHistory.Task) Then
                    Result.Task = SelectedHistory.Task
                End If
            End If
        End With
        Return Result
    End Function
    Private Sub OnItemSend(Item As System.Object, ByRef Cancel As Boolean) Handles Application.ItemSend
        Dim recipient As Outlook.Recipient = Nothing
        Dim recipients As Outlook.Recipients = Nothing
        Dim mail As Outlook.MailItem = TryCast(Item, Outlook.MailItem)
        If IsNothing(mail) Then Exit Sub 'Pl. delegált Task kiküldése esetén
        Dim SendAt = Now.AddSeconds(DeferredDeliveryTimeinSeconds)
        mail.DeferredDeliveryTime = SendAt
        'We can achieve this by using the MailItem.DeferredDeliveryTime [see: MailItem.DeferredDeliveryTime Property]. https://social.msdn.microsoft.com/Forums/vstudio/en-US/3e408d93-9f26-422f-a571-881d603578b7/programmatically-set-deferred-delivery?forum=vsto
        'mail.SaveSentMessageFolder = csak ugyanazon store-ban működik!

        Dim MatterUsed As MatterClass = Nothing
        Dim TaskOrFileRecorded As New ControlValuesGiven
        For Each formRegion In Globals.FormRegions
            If TypeOf formRegion IsNot HUI_TaskDocFormRegion Then Continue For
            Dim formRegionPresented As HUI_TaskDocFormRegion = CType(formRegion, HUI_TaskDocFormRegion)
            If Not formRegionPresented.OutlookItem.Equals(mail) OrElse formRegionPresented.Open = False Then Continue For
            If formRegionPresented.LastTabPage = HUI_TaskDocFormRegion.Tab.CreateNewTask Then
                TaskOrFileRecorded = RecordCreateNewTaskAsInformation(formRegionPresented)
            ElseIf formRegionPresented.LastTabPage = HUI_TaskDocFormRegion.Tab.FileToDocLibrary Then
                TaskOrFileRecorded = RecordFileSaveAsInformation(formRegionPresented)
            End If
        Next

        If Not IsNothing(MatterUsed) AndAlso Not String.IsNullOrWhiteSpace(MatterUsed.ObligatoryBCCAddresses) AndAlso Not mail.Sensitivity = Microsoft.Office.Interop.Outlook.OlSensitivity.olPersonal Then
            recipients = mail.Recipients
            Dim recipientlist As String() = Split(MatterUsed.ObligatoryBCCAddresses, ";")
            If recipientlist.Count > 0 Then
                For Each recs In recipientlist
                    recipient = recipients.Add(recs)
                    recipient.Type = Outlook.OlMailRecipientType.olBCC
                Next
                If Not recipients.ResolveAll Then
                    Dim strMsg = "Nem tudtam címezni az összes megadott Bcc címzettet. " & Environment.NewLine & "Így is elküldi?"
                    Dim res = MsgBox(strMsg, vbYesNo + vbDefaultButton1, "Bcc/Cc nem címezhető: " & MatterUsed.ObligatoryBCCAddresses)
                    If res = vbNo Then
                        Cancel = True
                    End If
                    'If Not IsNothing(recipient) Then Marshal.ReleaseComObject(recipient)
                    'If Not IsNothing(recipients) Then Marshal.ReleaseComObject(recipients)
                End If
            End If
        End If

        'SentItem esemény egy Outlook.Items esemény https://docs.microsoft.com/en-us/office/vba/api/outlook.items.itemadd
        SentItem = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Items
        'A SentMail folderbe ha bekerül egy mailItem, akkor lefut
        AddHandler SentItem.ItemAdd, Sub() AfterSent(SentItem, TaskOrFileRecorded)
    End Sub

    Private Sub AfterSent(item As System.Object, record As ControlValuesGiven)
        Dim mail As Outlook.MailItem = TryCast(item, Outlook.MailItem)
        If IsNothing(mail) Then Exit Sub
        'Ellenőrzi, hogy már le lett-e fűzve egyszer:
        If LastItemFiledGuid = mail.EntryID Then Exit Sub

        LastItemFiledGuid = mail.EntryID
        Dim SavedFileName = GetMailwAttachmentsIncludedAsFile(mail)
        Dim NewDocClass As DocClass
        If Not IsNothing(record.Task) Then
            'lefűzés taskhoz IS
            NewDocClass = HUI_TaskDocFormRegion.FillDocClassWithRecord(record, SavedFileName)
            HUI_TaskDocFormRegion.FileToDocLibraryAsFile(NewDocClass, record.Task.ID)
        Else
            'lefűzés csak docclasshoz
            NewDocClass = HUI_TaskDocFormRegion.FillDocClassWithRecord(record, SavedFileName)
            HUI_TaskDocFormRegion.FileToDocLibraryAsFile(NewDocClass)
        End If
        AddCategoryToMail(mail, SPHelper.SPHUI.FiledToSP)
        '#Itt át lehetne mozgatni másik mappába vagy törölni, de már le van fűzve
        '#Mi van, ha ki van kapcsolva a Sent Itemsbe lefűzés?
    End Sub

#End Region


End Class

