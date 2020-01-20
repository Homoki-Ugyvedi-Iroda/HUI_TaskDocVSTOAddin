Imports System.IO
Imports SPHelper.SPHUI
Imports System.Windows.Forms
Imports System.Threading.Tasks

Public Class ThisAddIn
    Property Connection As SPHelper.SPHUI
    Property TempPath As String

    Public Const UnknownFolderName = "/_Unknown/"
    Public Log As HPHelper.DebugTesting
    Public WithEvents TimerForRefresh As New Timer
    Public WithEvents TimerFirstSerialization As New Timer

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
    Public IsThisFirstStart As Boolean = False


    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Logstart()
        StartSp()
        SetAndVerifyTempPath()
        TimerForRefresh.Interval = 20 * 60 * 1000
        TimerForRefresh.Start()
        TimerFirstSerialization.Interval = 3 * 60 * 1000
        TimerFirstSerialization.Start()
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
    Private Sub StartSp()
        Me.Connection = New SPHelper.SPHUI()
        LoadCachedValues()
        LoadSpLookupValues()
    End Sub
    Private Sub SetAndVerifyTempPath()
        Globals.ThisAddIn.TempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName)
        If String.IsNullOrWhiteSpace(Me.TempPath) Then
            Globals.ThisAddIn.TempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName)
        End If
        If Not Directory.Exists(Me.TempPath) Then Directory.CreateDirectory(Me.TempPath)
        Globals.ThisAddIn.Log.logger.Info("SetAndVerifyTempPath finished")
    End Sub
    Private Sub TimerForRefresh_Tick(sender As Object, e As EventArgs) Handles TimerForRefresh.Tick
        RefreshSpLookupValues()
    End Sub
    Private Sub TimerFirstSerialization_Tick(sender As Object, e As EventArgs) Handles TimerFirstSerialization.Tick
        SerializeValues()
        TimerFirstSerialization.Stop()
    End Sub

    Public Async Sub LoadSpLookupValues()
        'deserialize-oltokat betölti, és beállítja az async-ot
        'Matter értékek: Id, Value, Active
        'Person értékek: adott matterhöz tartozó personok = Id, Value, Active, Matter
        'Users értékek: Id, loginname
        'Term értékek?
        'Task értékek?
        Await RefreshSpLookupValues()
    End Sub
    Public Async Function RefreshSpLookupValues() As Task
        Dim m As New DataLayer()
        Me.Connection.Connect(My.Settings.spUrl)
        If Me.Connection.Connected = False Then
            Globals.ThisAddIn.Log.logger.Info("No connection, no refresh")
            Exit Function
        End If
        Me.Connection.Users = Await m.GetAllUsers(Me.Connection)
        Me.Connection.Matters = Await m.GetAllMattersAsync(Me.Connection)
        Me.Connection.Persons = Await m.GetAllPersonsAsync(Me.Connection)
        Me.Connection.Tasks = Await m.GetAllTasksAsync(Me.Connection)
        Me.Connection.AllTerms = Await m.GetAllTermsAsync(Me.Connection)
        Me.Connection.Keywords = Await m.GetAllKeywords(Me.Connection)
        Me.Connection.InternalDocTypes = m.GetAllInternalDocTypes(Me.Connection)
        Me.Connection.WorkDocumentType = Await m.GetAllWorkDocTypesAsync(Me.Connection)
        Await LoadChoiceColumns(m)
        Globals.ThisAddIn.Log.logger.Info("RefreshSpLookupValues finished")
        If IsThisFirstStart = True Then
            SerializeValues()
        Else
            IsThisFirstStart = True
        End If
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
        If File.Exists(CWorkDocTypesPath) Then Me.Connection.WorkDocumentType = DeserializeWorkDocTypes(CWorkDocTypesPath)
        If File.Exists(CTermsPath) Then Me.Connection.AllTerms = DeserializeTerms(CTermsPath)
        If File.Exists(CInternalDocTypesPath) Then Me.Connection.InternalDocTypes = DeserializeTerms(CInternalDocTypesPath)
        If File.Exists(CKeywords) Then Me.Connection.Keywords = DeserializeTerms(CUsersPath)
        If File.Exists(CUsersPath) Then Me.Connection.Users = DeserializeUsers(CUsersPath)
        If File.Exists(CNonWorkDocTypesPath) Then Me.Connection.NonWorkDocTypes = DeserializeListofString(CNonWorkDocTypesPath)
        If File.Exists(CTaskTypesPath) Then Me.Connection.TaskTypes = DeserializeListofString(CTaskTypesPath)
        If File.Exists(CPrioritiesPath) Then Me.Connection.Priorities = DeserializeListofString(CPrioritiesPath)
        If File.Exists(CMattersPath) Then Me.Connection.Matters = DeserializeMatterClass(CMattersPath)
        If File.Exists(CPersonsPath) Then Me.Connection.Persons = DeserializePersonClass(CPersonsPath)
        If File.Exists(CTasksPath) Then Me.Connection.Tasks = DeserializeTaskClass(CTasksPath)
        Globals.ThisAddIn.Log.logger.Info("Values loaded from cache.")
    End Sub
    Private Sub SerializeValues()
        If Not Directory.Exists(CacheDirectory) Then
            Directory.CreateDirectory(CacheDirectory)
            Globals.ThisAddIn.Log.logger.Info("Cache directory created at " & CacheDirectory)
        End If
        Dim success As Byte = 0
        If Not IsNothing(Me.Connection.Keywords) AndAlso Me.Connection.Keywords.Count > 0 Then
            Dim succeeded = HPHelper.Serialize.Serialize(CKeywords, Me.Connection.Keywords)
            If succeeded Then success += 1
        End If
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
        If Not IsNothing(Me.Connection.AllTerms) AndAlso Me.Connection.AllTerms.Count > 0 Then
            Dim succeeded = HPHelper.Serialize.Serialize(CTermsPath, Me.Connection.AllTerms)
            If succeeded Then success += 1
        End If
        If Not IsNothing(Me.Connection.InternalDocTypes) AndAlso Me.Connection.InternalDocTypes.Count > 0 Then
            Dim succeeded = HPHelper.Serialize.Serialize(CInternalDocTypesPath, Me.Connection.InternalDocTypes)
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
        If Not IsNothing(Me.Connection.Matters) AndAlso Me.Connection.Matters.Count > 0 Then
            Dim succeeded = HPHelper.Serialize.Serialize(CMattersPath, Me.Connection.Matters)
            If succeeded Then success += 1
        End If
        Globals.ThisAddIn.Log.logger.Info("Items serialized from memory to file (" & success & ")")
    End Sub

End Class

