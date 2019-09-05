Imports System.IO
Imports SPHelper.SPHUI
Imports Microsoft.SharePoint.Client
Imports Microsoft.SharePoint.Client.Taxonomy
Imports System.Windows.Forms

Public Class ThisAddIn
    Property Connection As SPHelper.SPHUI
    Property LastTabPage As Integer
    Property ShowAllPartners As Boolean
    Property TempPath As String
    '>>ezeket átvinni új SPHUI classba
    'LoadSpLookupValues()

    Const DeferredDeliveryTimeinSeconds As Integer = 30
    Public Const UnknownFolderName = "/_Unknown/"
    Public LastTaskFiled As String
    Public Log As HPHelper.DebugTesting

    Private MyTimer As New System.Timers.Timer((DeferredDeliveryTimeinSeconds + 20) * 1000)
    Private WithEvents mySentItem As Outlook.Items
    Private LastItemFiledGuid As String

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Logstart()
        StartSp()
        SetAndVerifyTempPath()
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
            Connection = MySP
            Me.Connection = MySP
        Else
            Log.logger.Error("Nem érhető el a Sharepoint szerver, offline.")
        End If
    End Sub
    Private Sub SetAndVerifyTempPath()
        Globals.ThisAddIn.TempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName)
        If String.IsNullOrWhiteSpace(Globals.ThisAddIn.TempPath) Then
            Globals.ThisAddIn.TempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName)
        End If
        If Not Directory.Exists(Globals.ThisAddIn.TempPath) Then Directory.CreateDirectory(Globals.ThisAddIn.TempPath)
    End Sub
End Class
Public Class AddinSpHandler
    Inherits SPHelper.SPHUI
    'Property SPTimeZone As TimeZone
    'Property LCID As Integer
    'Property AllTerms As List(Of Term)
    'Property InternalDocTypes As List(Of Term)
    'Property Keywords As List(Of Term)
    'Property Matters As List(Of MatterClass)
    'Property NonWorkDocTypes As List(Of String)
    'Property Persons As List(Of PersonClass)
    'Property Priorities As List(Of String)
    'Property TaskTypes As List(Of String)
    'Property Users As List(Of User)
    'Property WorkDocumentType As List(Of WorkDocTypeClass)
    Property TaskTreeView As TreeView '>> kiegészíteni SPHelpert ezzel

End Class
