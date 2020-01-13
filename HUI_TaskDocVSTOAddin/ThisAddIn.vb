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

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Logstart()
        StartSp()
        SetAndVerifyTempPath()
        TimerForRefresh.Interval = 20 * 60 * 1000
        TimerForRefresh.Start()
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
        Dim MySP As New SPHelper.SPHUI()
        MySP.Connect(My.Settings.spUrl)
        If MySP.Connected = True Then
            Me.Connection = MySP
        Else
            Log.logger.Error("Nem érhető el a Sharepoint szerver, offline.")
        End If
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

    Public Async Sub LoadSpLookupValues()
        'deserialize-oltokat betölti, és beállítja az async-ot
        'Matter értékek: Id, Value, Active
        'Person értékek: adott matterhöz tartozó personok = Id, Value, Active, Matter
        'Users értékek: Id, loginname
        'Term értékek
        Await RefreshSpLookupValues()
        'Task!
    End Sub
    Public Async Function RefreshSpLookupValues() As Task
        Dim m As New DataLayer()
        Me.Connection.Users = Await m.GetAllUsers(Me.Connection)
        Me.Connection.Matters = Await m.GetAllMattersAsync(Me.Connection)
        Me.Connection.Persons = Await m.GetAllPersonsAsync(Me.Connection)
        Me.Connection.Tasks = Await m.GetAllTasksAsync(Me.Connection)
        Me.Connection.AllTerms = Await m.GetAllTermsAsync(Me.Connection)
        'Me.Connection.Keywords
        Me.Connection.InternalDocTypes = m.GetAllInternalDocTypes(Me.Connection)
        Me.Connection.WorkDocumentType = Await m.GetAllWorkDocTypesAsync(Me.Connection)
        Await LoadChoiceColumns(m)
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

End Class

