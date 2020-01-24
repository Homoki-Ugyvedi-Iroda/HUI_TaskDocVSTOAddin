<System.ComponentModel.ToolboxItemAttribute(False)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HUI_TaskDocFormRegion
    Inherits Microsoft.Office.Tools.Outlook.FormRegionBase

    Public Sub New(ByVal formRegion As Microsoft.Office.Interop.Outlook.FormRegion)
        MyBase.New(Globals.Factory, formRegion)
        Me.InitializeComponent()
    End Sub

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblConversationID = New System.Windows.Forms.Label()
        Me.btnFuzzLeElozmenyhez = New System.Windows.Forms.Button()
        Me.cbTasksFound = New System.Windows.Forms.ComboBox()
        Me.btnChooseOtherTask = New System.Windows.Forms.Button()
        Me.btnCsakCsatolmanyokLementese = New System.Windows.Forms.Button()
        Me.btnNewTask = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblConversationID
        '
        Me.lblConversationID.AutoSize = True
        Me.lblConversationID.Location = New System.Drawing.Point(3, 27)
        Me.lblConversationID.Name = "lblConversationID"
        Me.lblConversationID.Size = New System.Drawing.Size(98, 13)
        Me.lblConversationID.TabIndex = 0
        Me.lblConversationID.Text = "Valami label üzenet"
        '
        'btnFuzzLeElozmenyhez
        '
        Me.btnFuzzLeElozmenyhez.Location = New System.Drawing.Point(227, 27)
        Me.btnFuzzLeElozmenyhez.Name = "btnFuzzLeElozmenyhez"
        Me.btnFuzzLeElozmenyhez.Size = New System.Drawing.Size(116, 23)
        Me.btnFuzzLeElozmenyhez.TabIndex = 1
        Me.btnFuzzLeElozmenyhez.Text = "Fűzze le a fentihez"
        Me.btnFuzzLeElozmenyhez.UseVisualStyleBackColor = True
        '
        'cbTasksFound
        '
        Me.cbTasksFound.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbTasksFound.FormattingEnabled = True
        Me.cbTasksFound.Location = New System.Drawing.Point(3, 3)
        Me.cbTasksFound.Name = "cbTasksFound"
        Me.cbTasksFound.Size = New System.Drawing.Size(340, 21)
        Me.cbTasksFound.TabIndex = 2
        '
        'btnChooseOtherTask
        '
        Me.btnChooseOtherTask.Location = New System.Drawing.Point(227, 57)
        Me.btnChooseOtherTask.Name = "btnChooseOtherTask"
        Me.btnChooseOtherTask.Size = New System.Drawing.Size(116, 23)
        Me.btnChooseOtherTask.TabIndex = 3
        Me.btnChooseOtherTask.Text = "Másik feladathoz fűz"
        Me.btnChooseOtherTask.UseVisualStyleBackColor = True
        '
        'btnCsakCsatolmanyokLementese
        '
        Me.btnCsakCsatolmanyokLementese.Location = New System.Drawing.Point(107, 27)
        Me.btnCsakCsatolmanyokLementese.Name = "btnCsakCsatolmanyokLementese"
        Me.btnCsakCsatolmanyokLementese.Size = New System.Drawing.Size(116, 53)
        Me.btnCsakCsatolmanyokLementese.TabIndex = 4
        Me.btnCsakCsatolmanyokLementese.Text = "Csak csatolmányok lementése"
        Me.btnCsakCsatolmanyokLementese.UseVisualStyleBackColor = True
        '
        'btnNewTask
        '
        Me.btnNewTask.Location = New System.Drawing.Point(227, 86)
        Me.btnNewTask.Name = "btnNewTask"
        Me.btnNewTask.Size = New System.Drawing.Size(116, 23)
        Me.btnNewTask.TabIndex = 5
        Me.btnNewTask.Text = "Új feladathoz fűz"
        Me.btnNewTask.UseVisualStyleBackColor = True
        '
        'HUI_TaskDocFormRegion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btnNewTask)
        Me.Controls.Add(Me.btnCsakCsatolmanyokLementese)
        Me.Controls.Add(Me.btnChooseOtherTask)
        Me.Controls.Add(Me.cbTasksFound)
        Me.Controls.Add(Me.btnFuzzLeElozmenyhez)
        Me.Controls.Add(Me.lblConversationID)
        Me.Name = "HUI_TaskDocFormRegion"
        Me.Size = New System.Drawing.Size(357, 150)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    'NOTE: The following procedure is required by the Form Regions Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Shared Sub InitializeManifest(ByVal manifest As Microsoft.Office.Tools.Outlook.FormRegionManifest, ByVal factory As Microsoft.Office.Tools.Outlook.Factory)
        manifest.FormRegionName = "TaskDoc"
        manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Adjoining
        manifest.ShowReadingPane = False

    End Sub

    Friend WithEvents lblConversationID As Windows.Forms.Label

    Friend WithEvents btnFuzzLeElozmenyhez As Windows.Forms.Button
    Friend WithEvents cbTasksFound As Windows.Forms.ComboBox
    Friend WithEvents btnChooseOtherTask As Windows.Forms.Button

    Friend WithEvents btnCsakCsatolmanyokLementese As Windows.Forms.Button
    Friend WithEvents btnNewTask As Windows.Forms.Button

    Partial Public Class HUI_TaskDocFormRegionFactory
        Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory

        Public Event FormRegionInitializing As Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler

        Private _Manifest As Microsoft.Office.Tools.Outlook.FormRegionManifest


        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Public Sub New()
            Me._Manifest = Globals.Factory.CreateFormRegionManifest()
            HUI_TaskDocFormRegion.InitializeManifest(Me._Manifest, Globals.Factory)
        End Sub

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        ReadOnly Property Manifest() As Microsoft.Office.Tools.Outlook.FormRegionManifest Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.Manifest
            Get
                Return Me._Manifest
            End Get
        End Property

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Function CreateFormRegion(ByVal formRegion As Microsoft.Office.Interop.Outlook.FormRegion) As Microsoft.Office.Tools.Outlook.IFormRegion Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion
            Dim form As HUI_TaskDocFormRegion = New HUI_TaskDocFormRegion(formRegion)
            form.Factory = Me
            Return form
        End Function

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Function GetFormRegionStorage(ByVal outlookItem As Object, ByVal formRegionMode As Microsoft.Office.Interop.Outlook.OlFormRegionMode, ByVal formRegionSize As Microsoft.Office.Interop.Outlook.OlFormRegionSize) As Byte() Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage
            Throw New System.NotSupportedException()
        End Function

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Function IsDisplayedForItem(ByVal outlookItem As Object, ByVal formRegionMode As Microsoft.Office.Interop.Outlook.OlFormRegionMode, ByVal formRegionSize As Microsoft.Office.Interop.Outlook.OlFormRegionSize) As Boolean Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem
            Dim cancelArgs As Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, False)
            cancelArgs.Cancel = False
            RaiseEvent FormRegionInitializing(Me, cancelArgs)
            Return Not cancelArgs.Cancel
        End Function

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        ReadOnly Property Kind() As Microsoft.Office.Tools.Outlook.FormRegionKindConstants Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            Get
                Return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms
            End Get
        End Property
    End Class
End Class

Partial Class WindowFormRegionCollection

    Friend ReadOnly Property HUI_TaskDocFormRegion() As HUI_TaskDocFormRegion
        Get
            For Each Item As Object In Me
                If (TypeOf (Item) Is HUI_TaskDocFormRegion) Then
                    Return Item
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class