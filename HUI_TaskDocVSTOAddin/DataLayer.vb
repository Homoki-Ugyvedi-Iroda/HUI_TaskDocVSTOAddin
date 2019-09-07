Imports System.Threading.Tasks
Imports Microsoft.SharePoint.Client
Imports Microsoft.SharePoint.Client.Taxonomy
Imports SPHelper.SPHUI

Public Class DataLayer

    Public Async Function GetAllUsers(SphuiToUse As SPHelper.SPHUI) As Task(Of List(Of UserId))
        Return Await SphuiToUse.GetAllUsersAsync(SphuiToUse.Context, True, "Intranet HUI Members")
    End Function
    Public Async Function GetAllMattersAsync(SphuiToUse As SPHelper.SPHUI) As Task(Of List(Of MatterClass))
        Return Await SphuiToUse.GetAllMattersAsync(SphuiToUse.Context, SphuiToUse.Users)
    End Function
    Public Async Function GetAllPersonsAsync(SphuiToUse As SPHelper.SPHUI) As Task(Of List(Of PersonClass))
        Return Await SphuiToUse.GetAllPersonsAsyncTimesheet(SphuiToUse.Context, SphuiToUse.Matters)
    End Function
    Public Async Function GetAllTasksAsync(SphuiToUse As SPHelper.SPHUI) As Task(Of List(Of TaskClass))
        Return Await SphuiToUse.GetAllTasksForAddin(SphuiToUse.Context, SphuiToUse.Users, SphuiToUse.Matters, SphuiToUse.Persons)
    End Function
    Public Async Function GetAllTermsAsync(SphuiToUse As SPHelper.SPHUI) As Task(Of List(Of Term))
        Return Await SphuiToUse.GetAllTermsAsync(SphuiToUse.Context)
    End Function
    'Public Async Function GetAllKeywords(SphuiToUse As SPHelper.SPHUI) As Task(Of List(Of Term))
    '    Return Await SphuiToUse.GetAllTermsAsync(SphuiToUse.Context)
    'End Function
    Public Function GetAllInternalDocTypes(SphuiToUse As SPHelper.SPHUI) As List(Of Term)
        Return SphuiToUse.GetInternalDocTypes
    End Function
    Public Async Function GetAllWorkDocTypesAsync(SphuiToUse As SPHelper.SPHUI) As Task(Of List(Of WorkDocTypeClass))
        Return Await SphuiToUse.GetAllWorkDocTypesAsync(SphuiToUse.Context)
    End Function
    Public Async Function GetAllColumnChoiceTypesAsync(SphuiToUse As SPHelper.SPHUI) As Task(Of Dictionary(Of String, FieldMultiChoice))
        Return Await SphuiToUse.GetColumnValuesAsync(SphuiToUse.Context)
    End Function

End Class
