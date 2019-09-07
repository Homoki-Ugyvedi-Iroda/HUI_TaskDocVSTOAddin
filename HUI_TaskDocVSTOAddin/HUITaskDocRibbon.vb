Imports Microsoft.Office.Tools.Ribbon


Public Class HUITaskDocRibbon
    Private WithEvents mySentItem As Outlook.Items
    Private LastItemFiledGuid As String
    Const DeferredDeliveryTimeinSeconds As Integer = 30
    Public LastTaskFiled As String
    Private MyTimer As New System.Timers.Timer((DeferredDeliveryTimeinSeconds + 20) * 1000)

    Private Sub HUITaskDocRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

End Class
