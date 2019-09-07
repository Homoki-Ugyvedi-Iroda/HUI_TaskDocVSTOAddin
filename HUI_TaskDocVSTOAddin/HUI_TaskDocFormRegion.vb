Imports System.Windows.Forms

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

    'Occurs before the form region is displayed. 
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub HUI_TaskDocFormRegion_FormRegionShowing(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionShowing

    End Sub

    'Occurs when the form region is closed.   
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub HUI_TaskDocFormRegion_FormRegionClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionClosed

    End Sub

End Class
