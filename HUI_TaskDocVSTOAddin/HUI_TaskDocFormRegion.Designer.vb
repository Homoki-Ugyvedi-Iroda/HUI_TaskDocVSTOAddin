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
        Me.TableLayoutPanelTop = New System.Windows.Forms.TableLayoutPanel()
        Me.cbFileEmailOrAttachments = New System.Windows.Forms.ComboBox()
        Me.tbTitleFile = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabCreateTask = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanelCreateNewTask = New System.Windows.Forms.TableLayoutPanel()
        Me.cbSubTask = New System.Windows.Forms.CheckBox()
        Me.btnCreateNewTask = New System.Windows.Forms.Button()
        Me.btnExistingTaskChoiceAsTemplateForNewTask = New System.Windows.Forms.Button()
        Me.btnHistoryChosenAsTemplateForNewTask = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbTaskChosenHistoryNewTask = New System.Windows.Forms.ComboBox()
        Me.TabFileAsDoc = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanelFileToDocLibrary = New System.Windows.Forms.TableLayoutPanel()
        Me.btnChangePartnerOrder = New System.Windows.Forms.Button()
        Me.btnCreateFolderIfNotExisting = New System.Windows.Forms.Button()
        Me.lbTotalPartners = New System.Windows.Forms.ListBox()
        Me.btnExistingTaskChoiceAsFileTo_File = New System.Windows.Forms.Button()
        Me.btnFileToDocLibrary = New System.Windows.Forms.Button()
        Me.btnPartnerQryMailBody = New System.Windows.Forms.Button()
        Me.btnDeleteLastPartner = New System.Windows.Forms.Button()
        Me.btnAddPartnerFile = New System.Windows.Forms.Button()
        Me.cbPartner = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnOpenFolder = New System.Windows.Forms.Button()
        Me.tbPathToSaveTo = New System.Windows.Forms.TextBox()
        Me.cbMatter = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbFileHistory = New System.Windows.Forms.ComboBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TableLayoutPanelTop.SuspendLayout()
        Me.TabCreateTask.SuspendLayout()
        Me.TableLayoutPanelCreateNewTask.SuspendLayout()
        Me.TabFileAsDoc.SuspendLayout()
        Me.TableLayoutPanelFileToDocLibrary.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanelTop
        '
        Me.TableLayoutPanelTop.ColumnCount = 2
        Me.TableLayoutPanelTop.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanelTop.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanelTop.Controls.Add(Me.cbFileEmailOrAttachments, 0, 0)
        Me.TableLayoutPanelTop.Controls.Add(Me.TabControl1, 2, 1)
        Me.TableLayoutPanelTop.Controls.Add(Me.tbTitleFile, 1, 1)
        Me.TableLayoutPanelTop.Controls.Add(Me.Label1, 0, 1)
        Me.TableLayoutPanelTop.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanelTop.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanelTop.Name = "TableLayoutPanelTop"
        Me.TableLayoutPanelTop.RowCount = 3
        Me.TableLayoutPanelTop.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanelTop.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanelTop.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanelTop.Size = New System.Drawing.Size(719, 207)
        Me.TableLayoutPanelTop.TabIndex = 1
        '
        'cbFileEmailOrAttachments
        '
        Me.TableLayoutPanelTop.SetColumnSpan(Me.cbFileEmailOrAttachments, 2)
        Me.cbFileEmailOrAttachments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cbFileEmailOrAttachments.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbFileEmailOrAttachments.FormattingEnabled = True
        Me.cbFileEmailOrAttachments.Location = New System.Drawing.Point(3, 3)
        Me.cbFileEmailOrAttachments.Name = "cbFileEmailOrAttachments"
        Me.cbFileEmailOrAttachments.Size = New System.Drawing.Size(713, 21)
        Me.cbFileEmailOrAttachments.TabIndex = 0
        '
        'tbTitleFile
        '
        Me.tbTitleFile.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbTitleFile.Location = New System.Drawing.Point(73, 30)
        Me.tbTitleFile.Name = "tbTitleFile"
        Me.tbTitleFile.Size = New System.Drawing.Size(643, 20)
        Me.tbTitleFile.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Title (Name)"
        '
        'TabCreateTask
        '
        Me.TabCreateTask.BackColor = System.Drawing.Color.PaleGreen
        Me.TabCreateTask.Controls.Add(Me.TableLayoutPanelCreateNewTask)
        Me.TabCreateTask.Location = New System.Drawing.Point(4, 22)
        Me.TabCreateTask.Name = "TabCreateTask"
        Me.TabCreateTask.Size = New System.Drawing.Size(705, 118)
        Me.TabCreateTask.TabIndex = 2
        Me.TabCreateTask.Text = "Create New Task"
        '
        'TableLayoutPanelCreateNewTask
        '
        Me.TableLayoutPanelCreateNewTask.ColumnCount = 3
        Me.TableLayoutPanelCreateNewTask.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanelCreateNewTask.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanelCreateNewTask.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanelCreateNewTask.Controls.Add(Me.cbTaskChosenHistoryNewTask, 1, 0)
        Me.TableLayoutPanelCreateNewTask.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanelCreateNewTask.Controls.Add(Me.btnHistoryChosenAsTemplateForNewTask, 1, 1)
        Me.TableLayoutPanelCreateNewTask.Controls.Add(Me.btnExistingTaskChoiceAsTemplateForNewTask, 0, 2)
        Me.TableLayoutPanelCreateNewTask.Controls.Add(Me.btnCreateNewTask, 0, 3)
        Me.TableLayoutPanelCreateNewTask.Controls.Add(Me.cbSubTask, 2, 1)
        Me.TableLayoutPanelCreateNewTask.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanelCreateNewTask.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanelCreateNewTask.Name = "TableLayoutPanelCreateNewTask"
        Me.TableLayoutPanelCreateNewTask.RowCount = 4
        Me.TableLayoutPanelCreateNewTask.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanelCreateNewTask.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanelCreateNewTask.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanelCreateNewTask.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanelCreateNewTask.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanelCreateNewTask.Size = New System.Drawing.Size(705, 118)
        Me.TableLayoutPanelCreateNewTask.TabIndex = 0
        '
        'cbSubTask
        '
        Me.cbSubTask.AutoSize = True
        Me.cbSubTask.Location = New System.Drawing.Point(436, 30)
        Me.cbSubTask.Name = "cbSubTask"
        Me.cbSubTask.Size = New System.Drawing.Size(228, 17)
        Me.cbSubTask.TabIndex = 5
        Me.cbSubTask.Text = "Create subtask instead for the chosen task"
        Me.cbSubTask.UseVisualStyleBackColor = True
        '
        'btnCreateNewTask
        '
        Me.btnCreateNewTask.AutoSize = True
        Me.btnCreateNewTask.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnCreateNewTask.Location = New System.Drawing.Point(3, 88)
        Me.btnCreateNewTask.Name = "btnCreateNewTask"
        Me.btnCreateNewTask.Size = New System.Drawing.Size(156, 23)
        Me.btnCreateNewTask.TabIndex = 4
        Me.btnCreateNewTask.Text = "Create new task with this mail"
        Me.btnCreateNewTask.UseVisualStyleBackColor = True
        '
        'btnExistingTaskChoiceAsTemplateForNewTask
        '
        Me.btnExistingTaskChoiceAsTemplateForNewTask.AutoSize = True
        Me.btnExistingTaskChoiceAsTemplateForNewTask.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.TableLayoutPanelCreateNewTask.SetColumnSpan(Me.btnExistingTaskChoiceAsTemplateForNewTask, 2)
        Me.btnExistingTaskChoiceAsTemplateForNewTask.Location = New System.Drawing.Point(3, 59)
        Me.btnExistingTaskChoiceAsTemplateForNewTask.Name = "btnExistingTaskChoiceAsTemplateForNewTask"
        Me.btnExistingTaskChoiceAsTemplateForNewTask.Size = New System.Drawing.Size(259, 23)
        Me.btnExistingTaskChoiceAsTemplateForNewTask.TabIndex = 1
        Me.btnExistingTaskChoiceAsTemplateForNewTask.Text = "Choose other existing task as template for new task"
        Me.btnExistingTaskChoiceAsTemplateForNewTask.UseVisualStyleBackColor = True
        '
        'btnHistoryChosenAsTemplateForNewTask
        '
        Me.btnHistoryChosenAsTemplateForNewTask.AutoSize = True
        Me.btnHistoryChosenAsTemplateForNewTask.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnHistoryChosenAsTemplateForNewTask.Enabled = False
        Me.btnHistoryChosenAsTemplateForNewTask.Location = New System.Drawing.Point(165, 30)
        Me.btnHistoryChosenAsTemplateForNewTask.Name = "btnHistoryChosenAsTemplateForNewTask"
        Me.btnHistoryChosenAsTemplateForNewTask.Size = New System.Drawing.Size(191, 23)
        Me.btnHistoryChosenAsTemplateForNewTask.TabIndex = 3
        Me.btnHistoryChosenAsTemplateForNewTask.Text = "Use this task from history as template"
        Me.btnHistoryChosenAsTemplateForNewTask.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 5)
        Me.Label2.Margin = New System.Windows.Forms.Padding(5, 5, 5, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(144, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Use history tasks as template"
        '
        'cbTaskChosenHistoryNewTask
        '
        Me.TableLayoutPanelCreateNewTask.SetColumnSpan(Me.cbTaskChosenHistoryNewTask, 2)
        Me.cbTaskChosenHistoryNewTask.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cbTaskChosenHistoryNewTask.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbTaskChosenHistoryNewTask.FormattingEnabled = True
        Me.cbTaskChosenHistoryNewTask.Location = New System.Drawing.Point(165, 3)
        Me.cbTaskChosenHistoryNewTask.MaxDropDownItems = 20
        Me.cbTaskChosenHistoryNewTask.Name = "cbTaskChosenHistoryNewTask"
        Me.cbTaskChosenHistoryNewTask.Size = New System.Drawing.Size(537, 21)
        Me.cbTaskChosenHistoryNewTask.TabIndex = 0
        '
        'TabFileAsDoc
        '
        Me.TabFileAsDoc.BackColor = System.Drawing.Color.MistyRose
        Me.TabFileAsDoc.Controls.Add(Me.TableLayoutPanelFileToDocLibrary)
        Me.TabFileAsDoc.Location = New System.Drawing.Point(4, 22)
        Me.TabFileAsDoc.Name = "TabFileAsDoc"
        Me.TabFileAsDoc.Padding = New System.Windows.Forms.Padding(3)
        Me.TabFileAsDoc.Size = New System.Drawing.Size(705, 118)
        Me.TabFileAsDoc.TabIndex = 0
        Me.TabFileAsDoc.Text = "File to DocLibrary"
        '
        'TableLayoutPanelFileToDocLibrary
        '
        Me.TableLayoutPanelFileToDocLibrary.ColumnCount = 6
        Me.TableLayoutPanelFileToDocLibrary.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanelFileToDocLibrary.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 34.05066!))
        Me.TableLayoutPanelFileToDocLibrary.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 25.0!))
        Me.TableLayoutPanelFileToDocLibrary.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.26069!))
        Me.TableLayoutPanelFileToDocLibrary.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.34432!))
        Me.TableLayoutPanelFileToDocLibrary.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.34432!))
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.cbFileHistory, 1, 0)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.Label4, 0, 0)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.Label5, 0, 1)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.cbMatter, 1, 1)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.tbPathToSaveTo, 2, 1)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.btnOpenFolder, 4, 1)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.Label6, 0, 2)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.cbPartner, 1, 2)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.btnAddPartnerFile, 2, 2)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.btnDeleteLastPartner, 5, 2)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.btnPartnerQryMailBody, 5, 3)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.btnFileToDocLibrary, 1, 3)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.btnExistingTaskChoiceAsFileTo_File, 5, 0)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.lbTotalPartners, 3, 2)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.btnCreateFolderIfNotExisting, 5, 1)
        Me.TableLayoutPanelFileToDocLibrary.Controls.Add(Me.btnChangePartnerOrder, 2, 3)
        Me.TableLayoutPanelFileToDocLibrary.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanelFileToDocLibrary.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanelFileToDocLibrary.Name = "TableLayoutPanelFileToDocLibrary"
        Me.TableLayoutPanelFileToDocLibrary.RowCount = 4
        Me.TableLayoutPanelFileToDocLibrary.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanelFileToDocLibrary.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanelFileToDocLibrary.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanelFileToDocLibrary.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanelFileToDocLibrary.Size = New System.Drawing.Size(699, 112)
        Me.TableLayoutPanelFileToDocLibrary.TabIndex = 0
        '
        'btnChangePartnerOrder
        '
        Me.btnChangePartnerOrder.Font = New System.Drawing.Font("Segoe UI Symbol", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnChangePartnerOrder.Location = New System.Drawing.Point(302, 90)
        Me.btnChangePartnerOrder.Name = "btnChangePartnerOrder"
        Me.btnChangePartnerOrder.Size = New System.Drawing.Size(19, 23)
        Me.btnChangePartnerOrder.TabIndex = 21
        Me.btnChangePartnerOrder.Text = "↺"
        Me.btnChangePartnerOrder.UseVisualStyleBackColor = True
        '
        'btnCreateFolderIfNotExisting
        '
        Me.btnCreateFolderIfNotExisting.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCreateFolderIfNotExisting.AutoSize = True
        Me.btnCreateFolderIfNotExisting.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnCreateFolderIfNotExisting.Location = New System.Drawing.Point(607, 32)
        Me.btnCreateFolderIfNotExisting.Name = "btnCreateFolderIfNotExisting"
        Me.btnCreateFolderIfNotExisting.Size = New System.Drawing.Size(89, 23)
        Me.btnCreateFolderIfNotExisting.TabIndex = 20
        Me.btnCreateFolderIfNotExisting.Text = "Create Folder"
        Me.btnCreateFolderIfNotExisting.UseVisualStyleBackColor = True
        '
        'lbTotalPartners
        '
        Me.lbTotalPartners.BackColor = System.Drawing.SystemColors.ControlLight
        Me.TableLayoutPanelFileToDocLibrary.SetColumnSpan(Me.lbTotalPartners, 2)
        Me.lbTotalPartners.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbTotalPartners.FormattingEnabled = True
        Me.lbTotalPartners.Location = New System.Drawing.Point(327, 61)
        Me.lbTotalPartners.Name = "lbTotalPartners"
        Me.TableLayoutPanelFileToDocLibrary.SetRowSpan(Me.lbTotalPartners, 2)
        Me.lbTotalPartners.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.lbTotalPartners.Size = New System.Drawing.Size(274, 52)
        Me.lbTotalPartners.TabIndex = 19
        '
        'btnExistingTaskChoiceAsFileTo_File
        '
        Me.btnExistingTaskChoiceAsFileTo_File.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExistingTaskChoiceAsFileTo_File.AutoSize = True
        Me.btnExistingTaskChoiceAsFileTo_File.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnExistingTaskChoiceAsFileTo_File.Location = New System.Drawing.Point(607, 3)
        Me.btnExistingTaskChoiceAsFileTo_File.Name = "btnExistingTaskChoiceAsFileTo_File"
        Me.btnExistingTaskChoiceAsFileTo_File.Size = New System.Drawing.Size(89, 23)
        Me.btnExistingTaskChoiceAsFileTo_File.TabIndex = 18
        Me.btnExistingTaskChoiceAsFileTo_File.Text = "Choose Task"
        Me.btnExistingTaskChoiceAsFileTo_File.UseVisualStyleBackColor = True
        '
        'btnFileToDocLibrary
        '
        Me.btnFileToDocLibrary.Location = New System.Drawing.Point(109, 90)
        Me.btnFileToDocLibrary.Name = "btnFileToDocLibrary"
        Me.btnFileToDocLibrary.Size = New System.Drawing.Size(75, 23)
        Me.btnFileToDocLibrary.TabIndex = 17
        Me.btnFileToDocLibrary.Text = "File"
        Me.btnFileToDocLibrary.UseVisualStyleBackColor = True
        '
        'btnPartnerQryMailBody
        '
        Me.btnPartnerQryMailBody.AutoSize = True
        Me.btnPartnerQryMailBody.Location = New System.Drawing.Point(607, 90)
        Me.btnPartnerQryMailBody.Name = "btnPartnerQryMailBody"
        Me.btnPartnerQryMailBody.Size = New System.Drawing.Size(89, 23)
        Me.btnPartnerQryMailBody.TabIndex = 16
        Me.btnPartnerQryMailBody.Text = "PartnerQryMail"
        Me.btnPartnerQryMailBody.UseVisualStyleBackColor = True
        '
        'btnDeleteLastPartner
        '
        Me.btnDeleteLastPartner.Location = New System.Drawing.Point(607, 61)
        Me.btnDeleteLastPartner.Name = "btnDeleteLastPartner"
        Me.btnDeleteLastPartner.Size = New System.Drawing.Size(22, 23)
        Me.btnDeleteLastPartner.TabIndex = 15
        Me.btnDeleteLastPartner.Text = ""
        Me.btnDeleteLastPartner.UseVisualStyleBackColor = True
        '
        'btnAddPartnerFile
        '
        Me.btnAddPartnerFile.Location = New System.Drawing.Point(302, 61)
        Me.btnAddPartnerFile.Name = "btnAddPartnerFile"
        Me.btnAddPartnerFile.Size = New System.Drawing.Size(19, 23)
        Me.btnAddPartnerFile.TabIndex = 13
        Me.btnAddPartnerFile.Text = "»"
        Me.btnAddPartnerFile.UseVisualStyleBackColor = True
        '
        'cbPartner
        '
        Me.cbPartner.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbPartner.FormattingEnabled = True
        Me.cbPartner.Location = New System.Drawing.Point(109, 61)
        Me.cbPartner.Name = "cbPartner"
        Me.cbPartner.Size = New System.Drawing.Size(187, 21)
        Me.cbPartner.TabIndex = 12
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(5, 63)
        Me.Label6.Margin = New System.Windows.Forms.Padding(5, 5, 5, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(89, 13)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Partners involved"
        '
        'btnOpenFolder
        '
        Me.btnOpenFolder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOpenFolder.Location = New System.Drawing.Point(515, 32)
        Me.btnOpenFolder.Name = "btnOpenFolder"
        Me.btnOpenFolder.Size = New System.Drawing.Size(86, 23)
        Me.btnOpenFolder.TabIndex = 9
        Me.btnOpenFolder.Text = "Open Folder"
        Me.btnOpenFolder.UseVisualStyleBackColor = True
        '
        'tbPathToSaveTo
        '
        Me.TableLayoutPanelFileToDocLibrary.SetColumnSpan(Me.tbPathToSaveTo, 2)
        Me.tbPathToSaveTo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbPathToSaveTo.Location = New System.Drawing.Point(302, 32)
        Me.tbPathToSaveTo.Name = "tbPathToSaveTo"
        Me.tbPathToSaveTo.Size = New System.Drawing.Size(207, 20)
        Me.tbPathToSaveTo.TabIndex = 8
        '
        'cbMatter
        '
        Me.cbMatter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbMatter.FormattingEnabled = True
        Me.cbMatter.Location = New System.Drawing.Point(109, 32)
        Me.cbMatter.Name = "cbMatter"
        Me.cbMatter.Size = New System.Drawing.Size(187, 21)
        Me.cbMatter.TabIndex = 7
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 34)
        Me.Label5.Margin = New System.Windows.Forms.Padding(5, 5, 5, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(44, 13)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Matter#"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(5, 5)
        Me.Label4.Margin = New System.Windows.Forms.Padding(5, 5, 0, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(101, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "History for metadata"
        '
        'cbFileHistory
        '
        Me.TableLayoutPanelFileToDocLibrary.SetColumnSpan(Me.cbFileHistory, 3)
        Me.cbFileHistory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cbFileHistory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbFileHistory.FormattingEnabled = True
        Me.cbFileHistory.Location = New System.Drawing.Point(109, 3)
        Me.cbFileHistory.MaxDropDownItems = 20
        Me.cbFileHistory.Name = "cbFileHistory"
        Me.cbFileHistory.Size = New System.Drawing.Size(400, 21)
        Me.cbFileHistory.TabIndex = 5
        '
        'TabControl1
        '
        Me.TableLayoutPanelTop.SetColumnSpan(Me.TabControl1, 2)
        Me.TabControl1.Controls.Add(Me.TabFileAsDoc)
        Me.TabControl1.Controls.Add(Me.TabCreateTask)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(3, 60)
        Me.TabControl1.MaximumSize = New System.Drawing.Size(1920, 200)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(713, 144)
        Me.TabControl1.TabIndex = 0
        '
        'HUI_TaskDocFormRegion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanelTop)
        Me.Name = "HUI_TaskDocFormRegion"
        Me.Size = New System.Drawing.Size(719, 207)
        Me.TableLayoutPanelTop.ResumeLayout(False)
        Me.TableLayoutPanelTop.PerformLayout()
        Me.TabCreateTask.ResumeLayout(False)
        Me.TableLayoutPanelCreateNewTask.ResumeLayout(False)
        Me.TableLayoutPanelCreateNewTask.PerformLayout()
        Me.TabFileAsDoc.ResumeLayout(False)
        Me.TableLayoutPanelFileToDocLibrary.ResumeLayout(False)
        Me.TableLayoutPanelFileToDocLibrary.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

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

    Friend WithEvents TableLayoutPanelTop As Windows.Forms.TableLayoutPanel

    Friend WithEvents cbFileEmailOrAttachments As Windows.Forms.ComboBox

    Friend WithEvents tbTitleFile As Windows.Forms.TextBox

    Friend WithEvents Label1 As Windows.Forms.Label
    Public WithEvents TabControl1 As Windows.Forms.TabControl

    Public WithEvents TabFileAsDoc As Windows.Forms.TabPage

    Friend WithEvents TableLayoutPanelFileToDocLibrary As Windows.Forms.TableLayoutPanel

    Friend WithEvents cbFileHistory As Windows.Forms.ComboBox

    Friend WithEvents Label4 As Windows.Forms.Label

    Friend WithEvents Label5 As Windows.Forms.Label

    Friend WithEvents cbMatter As Windows.Forms.ComboBox

    Friend WithEvents tbPathToSaveTo As Windows.Forms.TextBox

    Friend WithEvents btnOpenFolder As Windows.Forms.Button

    Friend WithEvents Label6 As Windows.Forms.Label

    Friend WithEvents cbPartner As Windows.Forms.ComboBox

    Friend WithEvents btnAddPartnerFile As Windows.Forms.Button

    Friend WithEvents btnDeleteLastPartner As Windows.Forms.Button

    Friend WithEvents btnPartnerQryMailBody As Windows.Forms.Button

    Friend WithEvents btnFileToDocLibrary As Windows.Forms.Button

    Friend WithEvents btnExistingTaskChoiceAsFileTo_File As Windows.Forms.Button

    Friend WithEvents lbTotalPartners As Windows.Forms.ListBox

    Friend WithEvents btnCreateFolderIfNotExisting As Windows.Forms.Button

    Friend WithEvents btnChangePartnerOrder As Windows.Forms.Button

    Public WithEvents TabCreateTask As Windows.Forms.TabPage

    Friend WithEvents TableLayoutPanelCreateNewTask As Windows.Forms.TableLayoutPanel

    Friend WithEvents cbTaskChosenHistoryNewTask As Windows.Forms.ComboBox

    Friend WithEvents Label2 As Windows.Forms.Label

    Friend WithEvents btnHistoryChosenAsTemplateForNewTask As Windows.Forms.Button

    Friend WithEvents btnExistingTaskChoiceAsTemplateForNewTask As Windows.Forms.Button

    Friend WithEvents btnCreateNewTask As Windows.Forms.Button

    Friend WithEvents cbSubTask As Windows.Forms.CheckBox

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