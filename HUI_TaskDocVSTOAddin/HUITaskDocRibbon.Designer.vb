Partial Class HUITaskDocRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Task = Me.Factory.CreateRibbonGroup
        Me.btnReplyFromTask = Me.Factory.CreateRibbonButton
        Me.btnReplyToThisEmailFromTask = Me.Factory.CreateRibbonButton
        Me.cbxWithCompleted = Me.Factory.CreateRibbonCheckBox
        Me.Tab1.SuspendLayout()
        Me.Task.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Task)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Task
        '
        Me.Task.Items.Add(Me.btnReplyFromTask)
        Me.Task.Items.Add(Me.btnReplyToThisEmailFromTask)
        Me.Task.Items.Add(Me.cbxWithCompleted)
        Me.Task.Label = "Task"
        Me.Task.Name = "Task"
        '
        'btnReplyFromTask
        '
        Me.btnReplyFromTask.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnReplyFromTask.Image = Global.HUI_TaskDocVSTOAddin.My.Resources.Resources.TaskList_48x
        Me.btnReplyFromTask.Label = "Create reply e-mail from task"
        Me.btnReplyFromTask.Name = "btnReplyFromTask"
        Me.btnReplyFromTask.ShowImage = True
        '
        'btnReplyToThisEmailFromTask
        '
        Me.btnReplyToThisEmailFromTask.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnReplyToThisEmailFromTask.Image = Global.HUI_TaskDocVSTOAddin.My.Resources.Resources.InsertMark_12990_32
        Me.btnReplyToThisEmailFromTask.Label = "Reply to selected email from task"
        Me.btnReplyToThisEmailFromTask.Name = "btnReplyToThisEmailFromTask"
        Me.btnReplyToThisEmailFromTask.ShowImage = True
        '
        'cbxWithCompleted
        '
        Me.cbxWithCompleted.Checked = True
        Me.cbxWithCompleted.Label = "Complete task with reply"
        Me.cbxWithCompleted.Name = "cbxWithCompleted"
        '
        'HUITaskDocRibbon
        '
        Me.Name = "HUITaskDocRibbon"
        Me.RibbonType = "Microsoft.Outlook.Mail.Read"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Task.ResumeLayout(False)
        Me.Task.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Task As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnReplyFromTask As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnReplyToThisEmailFromTask As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents cbxWithCompleted As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property HUITaskDocRibbon() As HUITaskDocRibbon
        Get
            Return Me.GetRibbon(Of HUITaskDocRibbon)()
        End Get
    End Property
End Class
