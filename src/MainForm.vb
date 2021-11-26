
Imports System.Text.RegularExpressions
Imports System.Threading

Imports VB6 = Microsoft.VisualBasic

Public Class MainForm
    Inherits Form

#Region " Designer "

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    Friend WithEvents tv As System.Windows.Forms.TreeView
    Friend WithEvents chValue As System.Windows.Forms.ColumnHeader
    Friend WithEvents chName As System.Windows.Forms.ColumnHeader
    Friend WithEvents lv As System.Windows.Forms.ListView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lGuid As System.Windows.Forms.Label
    Friend WithEvents tsMain As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbRegedit As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbFolder As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbDir As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbAll As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbAllObjects As System.Windows.Forms.ToolStripButton
    Friend WithEvents cb As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents spMain As System.Windows.Forms.SplitContainer
    Friend WithEvents gbGuid As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tv = New System.Windows.Forms.TreeView()
        Me.lv = New System.Windows.Forms.ListView()
        Me.chName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.chValue = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.gbGuid = New System.Windows.Forms.GroupBox()
        Me.lGuid = New System.Windows.Forms.Label()
        Me.tsMain = New System.Windows.Forms.ToolStrip()
        Me.tsbAll = New System.Windows.Forms.ToolStripButton()
        Me.tsbAllObjects = New System.Windows.Forms.ToolStripButton()
        Me.tsbDir = New System.Windows.Forms.ToolStripButton()
        Me.tsbFolder = New System.Windows.Forms.ToolStripButton()
        Me.tsbRegedit = New System.Windows.Forms.ToolStripButton()
        Me.cb = New System.Windows.Forms.ToolStripComboBox()
        Me.spMain = New System.Windows.Forms.SplitContainer()
        Me.gbGuid.SuspendLayout()
        Me.tsMain.SuspendLayout()
        CType(Me.spMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.spMain.Panel1.SuspendLayout()
        Me.spMain.Panel2.SuspendLayout()
        Me.spMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'tv
        '
        Me.tv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tv.HideSelection = False
        Me.tv.Location = New System.Drawing.Point(12, 0)
        Me.tv.Name = "tv"
        Me.tv.Size = New System.Drawing.Size(859, 1209)
        Me.tv.TabIndex = 2
        '
        'lv
        '
        Me.lv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lv.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chName, Me.chValue})
        Me.lv.FullRowSelect = True
        Me.lv.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lv.HideSelection = False
        Me.lv.Location = New System.Drawing.Point(0, 0)
        Me.lv.Margin = New System.Windows.Forms.Padding(3, 3, 12, 3)
        Me.lv.MultiSelect = False
        Me.lv.Name = "lv"
        Me.lv.Size = New System.Drawing.Size(1203, 905)
        Me.lv.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lv.TabIndex = 4
        Me.lv.UseCompatibleStateImageBehavior = False
        Me.lv.View = System.Windows.Forms.View.Details
        '
        'chName
        '
        Me.chName.Text = "Name"
        Me.chName.Width = 157
        '
        'chValue
        '
        Me.chValue.Text = "Value"
        Me.chValue.Width = 470
        '
        'gbGuid
        '
        Me.gbGuid.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbGuid.Controls.Add(Me.lGuid)
        Me.gbGuid.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.gbGuid.Location = New System.Drawing.Point(7, 910)
        Me.gbGuid.Name = "gbGuid"
        Me.gbGuid.Size = New System.Drawing.Size(1195, 300)
        Me.gbGuid.TabIndex = 3
        Me.gbGuid.TabStop = False
        Me.gbGuid.Text = "Guid"
        '
        'lGuid
        '
        Me.lGuid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lGuid.Location = New System.Drawing.Point(1, 54)
        Me.lGuid.Name = "lGuid"
        Me.lGuid.Size = New System.Drawing.Size(1183, 216)
        Me.lGuid.TabIndex = 0
        '
        'tsMain
        '
        Me.tsMain.AutoSize = False
        Me.tsMain.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.tsMain.ImageScalingSize = New System.Drawing.Size(48, 48)
        Me.tsMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbAll, Me.tsbAllObjects, Me.tsbDir, Me.tsbFolder, Me.tsbRegedit, Me.cb})
        Me.tsMain.Location = New System.Drawing.Point(0, 0)
        Me.tsMain.Name = "tsMain"
        Me.tsMain.Padding = New System.Windows.Forms.Padding(3, 1, 1, 0)
        Me.tsMain.Size = New System.Drawing.Size(2101, 80)
        Me.tsMain.TabIndex = 24
        '
        'tsbAll
        '
        Me.tsbAll.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbAll.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbAll.Name = "tsbAll"
        Me.tsbAll.Size = New System.Drawing.Size(185, 70)
        Me.tsbAll.Text = "  All Files  "
        Me.tsbAll.ToolTipText = "Opens all keys related to all files"
        '
        'tsbAllObjects
        '
        Me.tsbAllObjects.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbAllObjects.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbAllObjects.Name = "tsbAllObjects"
        Me.tsbAllObjects.Size = New System.Drawing.Size(235, 70)
        Me.tsbAllObjects.Text = "  All Objects  "
        Me.tsbAllObjects.ToolTipText = "Opens all keys related to all file system objects"
        '
        'tsbDir
        '
        Me.tsbDir.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbDir.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbDir.Name = "tsbDir"
        Me.tsbDir.Size = New System.Drawing.Size(210, 70)
        Me.tsbDir.Text = "  Directory  "
        Me.tsbDir.ToolTipText = "Opens all keys related to directories"
        '
        'tsbFolder
        '
        Me.tsbFolder.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbFolder.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbFolder.Name = "tsbFolder"
        Me.tsbFolder.Size = New System.Drawing.Size(165, 70)
        Me.tsbFolder.Text = "  Folder  "
        Me.tsbFolder.ToolTipText = "Opens all keys related to folders"
        '
        'tsbRegedit
        '
        Me.tsbRegedit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbRegedit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbRegedit.Name = "tsbRegedit"
        Me.tsbRegedit.Size = New System.Drawing.Size(186, 70)
        Me.tsbRegedit.Text = "  Regedit  "
        Me.tsbRegedit.ToolTipText = "Opens Regedit with the selected registry key"
        '
        'cb
        '
        Me.cb.FlatStyle = System.Windows.Forms.FlatStyle.Standard
        Me.cb.Name = "cb"
        Me.cb.Size = New System.Drawing.Size(300, 79)
        '
        'spMain
        '
        Me.spMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.spMain.Location = New System.Drawing.Point(0, 80)
        Me.spMain.Margin = New System.Windows.Forms.Padding(3, 0, 3, 3)
        Me.spMain.Name = "spMain"
        '
        'spMain.Panel1
        '
        Me.spMain.Panel1.Controls.Add(Me.tv)
        Me.spMain.Panel1.Padding = New System.Windows.Forms.Padding(12, 0, 8, 12)
        '
        'spMain.Panel2
        '
        Me.spMain.Panel2.Controls.Add(Me.gbGuid)
        Me.spMain.Panel2.Controls.Add(Me.lv)
        Me.spMain.Panel2.Padding = New System.Windows.Forms.Padding(0, 0, 12, 0)
        Me.spMain.Size = New System.Drawing.Size(2101, 1221)
        Me.spMain.SplitterDistance = 879
        Me.spMain.TabIndex = 25
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(288.0!, 288.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(2101, 1301)
        Me.Controls.Add(Me.spMain)
        Me.Controls.Add(Me.tsMain)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "MainForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "File Type Registry Viewer"
        Me.gbGuid.ResumeLayout(False)
        Me.tsMain.ResumeLayout(False)
        Me.tsMain.PerformLayout()
        Me.spMain.Panel1.ResumeLayout(False)
        Me.spMain.Panel2.ResumeLayout(False)
        CType(Me.spMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.spMain.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub New()
        MyBase.New()
        InitializeComponent()

        cb.Text = VB6.GetSetting(Application.ProductName, "Misc", "Type", ".txt")
        cb.Width = CInt(Font.Size * 10)

        ActiveControl = MainMenuStrip

        Native.SetWindowTheme(tv.Handle, "explorer", Nothing)
        Native.SetWindowTheme(lv.Handle, "explorer", Nothing)
    End Sub

    <STAThread()>
    Public Shared Sub Main()
        AddHandler Application.ThreadException, AddressOf ThreadException
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New MainForm)
    End Sub

    Public Shared Sub ThreadException(ByVal sender As Object, ByVal e As ThreadExceptionEventArgs)
        Msg(e.Exception.Message.ToString, MessageBoxIcon.Error)
    End Sub

    Private LastSelectedPath As String

    Private Sub tv_AfterSelect() Handles tv.AfterSelect
        PopulateValues()
        UpdateGUID()
        LastSelectedPath = tv.SelectedNode.FullPath
    End Sub

    Private Sub PopulateValues()
        lv.Items.Clear()

        If Not tv.SelectedNode Is Nothing Then
            Dim n = DirectCast(tv.SelectedNode, ExtendedTreeNode)

            Try
                For Each i In n.Key.GetValueNames
                    Dim value = n.Key.GetValue(i)

                    If Not value Is Nothing Then
                        lv.Items.Add(New ListViewItem(New String() {i, value.ToString}))
                    End If
                Next
            Catch
            End Try
        End If

        If lv.Items.Count > 0 Then
            lv.Items(0).Selected = True
        End If
    End Sub

    Private Sub UpdateGUID()
        Dim re As New Regex("^.*(\{\w{8}-\w{4}-\w{4}-\w{4}-\w{12}\})$")
        Dim value As String = Nothing
        Dim match As Match = Nothing

        If Not tv.SelectedNode Is Nothing Then
            match = re.Match(tv.SelectedNode.Text) 're.Match(lv.SelectedItems(0).SubItems(1).Text)

            If match.Success Then
                value = match.Groups(1).Value
            ElseIf Not lv.SelectedItems Is Nothing AndAlso lv.SelectedItems.Count > 0 Then
                match = re.Match(lv.SelectedItems(0).Text)

                If match.Success Then
                    value = match.Groups(1).Value
                Else
                    match = re.Match(lv.SelectedItems(0).SubItems(1).Text)

                    If match.Success Then
                        value = match.Groups(1).Value
                    End If
                End If
            End If
        End If

        If Not value Is Nothing Then
            Dim guidKey As RegistryKey = Nothing
            Dim guidKeyTemp = Registry.ClassesRoot.OpenSubKey("CLSID\" + value)

            If Not guidKeyTemp Is Nothing Then
                guidKey = guidKeyTemp
            Else
                guidKeyTemp = Registry.ClassesRoot.OpenSubKey("Wow6432Node\CLSID\" + value)

                If Not guidKeyTemp Is Nothing Then
                    guidKey = guidKeyTemp
                End If
            End If

            Using guidKey
                If Not guidKey Is Nothing Then
                    Using inprocServer32Key = guidKey.OpenSubKey("InprocServer32")
                        If Not inprocServer32Key Is Nothing Then
                            Dim fp = DirectCast(inprocServer32Key.GetValue(Nothing, ""), String)

                            If IO.File.Exists(fp) Then
                                lGuid.Text = ""
                                Dim append = Sub(temp As String) If temp <> "" Then lGuid.Text += temp + CrLf

                                If guidKey.Name.Contains("Wow6432Node") Then
                                    append("only x86!")
                                End If

                                append(DirectCast(guidKey.GetValue(Nothing, ""), String))
                                Dim i = FileVersionInfo.GetVersionInfo(fp)
                                append(i.ProductName)
                                append(i.FileDescription)
                                append(fp)
                                Exit Sub
                            End If
                        End If
                    End Using
                End If
            End Using
        End If

        lGuid.Text = ""
    End Sub

    Private Sub lv_SelectedIndexChanged() Handles lv.SelectedIndexChanged
        UpdateGUID()
    End Sub

    Private Sub tsbRegedit_Click() Handles tsbRegedit.Click
        ShowRegedit()
    End Sub

    Private Sub ShowRegedit()
        Dim n = DirectCast(tv.SelectedNode, ExtendedTreeNode)

        If Not n Is Nothing Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit",
                              "LastKey", n.Key.GetNativeName)
        End If

        Try 'cancelling UAC prompt throws exception
            Process.Start("regedit")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub tsbFolder_Click() Handles tsbFolder.Click
        cb.Text = "Folder"
    End Sub

    Private Sub tsbDir_Click() Handles tsbDir.Click
        cb.Text = "Directory"
    End Sub

    Private Sub tsbAll_Click() Handles tsbAll.Click
        cb.Text = "*"
    End Sub

    Private Sub tsbAllObjects_Click() Handles tsbAllObjects.Click
        cb.Text = "AllFilesystemObjects"
    End Sub

    Private Sub MainForm_Activated() Handles Me.Activated
        cb_TextChanged()

        If Not LastSelectedPath Is Nothing Then
            SelectLastNode(tv.Nodes)
        End If
    End Sub

    Private Sub SelectLastNode(ByVal nodes As TreeNodeCollection)
        For Each i As TreeNode In nodes
            If i.FullPath = LastSelectedPath Then
                tv.SelectedNode = i
                Exit For
            End If

            SelectLastNode(i.Nodes)
        Next
    End Sub

    Private Sub MainForm_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles Me.FormClosed
        VB6.SaveSetting(Application.ProductName, "Misc", "Type", cb.Text)
    End Sub

    Private Sub MainForm_Shown() Handles Me.Shown
        Refresh()

        Dim l As New List(Of String)

        For Each i In Registry.ClassesRoot.GetSubKeyNames
            If i.StartsWith(".") Then
                l.Add(i)
            End If
        Next

        cb.Items.AddRange(l.ToArray)
        cb.AutoCompleteSource = AutoCompleteSource.ListItems
        cb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
    End Sub

    Private Sub cb_TextChanged() Handles cb.TextChanged
        If Not Visible Then
            Exit Sub
        End If

        tv.BeginUpdate()
        tv.Nodes.Clear()
        lv.Items.Clear()

        Dim HKCR As New Key("Classes Root", Registry.ClassesRoot.Name)
        Dim extkey = HKCR.GetSubKey(cb.Text)

        If Not extkey Is Nothing Then
            tv.Nodes.Add(HKCR.TreeNode)
            Key.BuildSubNoded(extkey)
            Dim defaultValue = CStr(extkey.GetValue(Nothing))

            If Not defaultValue Is Nothing Then
                Dim k = HKCR.GetSubKey(defaultValue)

                If Not k Is Nothing Then
                    k.BuildSubNoded()
                End If
            End If
        End If

        Dim HKCU As New Key("Current User", Registry.CurrentUser.Name)

        Dim explorerKey = HKCU.GetSubKey(
            "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" + cb.Text)

        If Not explorerKey Is Nothing Then
            explorerKey.BuildSubNoded()
            tv.Nodes(tv.Nodes.Add(HKCU.TreeNode)).ExpandAll()

            For Each i In explorerKey.GetValueNames
                If i = "Progid" Then
                    Dim value = CStr(explorerKey.GetValue(i))

                    If Not value Is Nothing Then
                        Dim typeKey = HKCR.GetSubKey(value)

                        If Not typeKey Is Nothing Then
                            typeKey.BuildSubNoded()
                        End If
                    End If
                End If
            Next

            Dim uc = explorerKey.GetSubKey("UserChoice")

            If Not uc Is Nothing Then
                Dim progID = CStr(uc.GetValue("Progid"))

                If Not progID Is Nothing Then
                    Dim k = HKCR.GetSubKey(progID)

                    If Not k Is Nothing Then
                        k.BuildSubNoded()
                    End If
                End If
            End If
        End If

        Dim HKLM As New Key("Local Machine", Registry.LocalMachine.Name)

        Dim propHandlerKey = HKLM.GetSubKey(
            "SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers\" + cb.Text)

        If Not propHandlerKey Is Nothing Then
            propHandlerKey.BuildSubNoded()
            tv.Nodes(tv.Nodes.Add(HKLM.TreeNode)).ExpandAll()
        End If

        Dim systemExtKey = HKCR.GetSubKey("SystemFileAssociations\" + cb.Text)

        If Not systemExtKey Is Nothing Then
            systemExtKey.BuildSubNoded()
        End If

        If Not extkey Is Nothing Then
            For Each i In extkey.GetValueNames
                If i = "PerceivedType" Then
                    Dim t = DirectCast(extkey.GetValue(i), String)

                    If Not t Is Nothing Then
                        Dim systemTypeKey = HKCR.GetSubKey(
                            "SystemFileAssociations\" + t)

                        If Not systemTypeKey Is Nothing Then
                            systemTypeKey.BuildSubNoded()
                        End If
                    End If
                End If
            Next
        End If

        If tv.Nodes.Count > 0 Then
            tv.Nodes(0).Expand()

            If tv.Nodes.Count = 1 AndAlso tv.Nodes(0).Nodes.Count > 0 Then
                tv.Nodes(0).Nodes(0).Expand()
            End If
        End If

        tv.EndUpdate()

        tsbFolder.Checked = False
        tsbDir.Checked = False
        tsbAll.Checked = False
        tsbAllObjects.Checked = False

        Select Case cb.Text
            Case "Folder"
                tsbFolder.Checked = True
            Case "Directory"
                tsbDir.Checked = True
            Case "*"
                tsbAll.Checked = True
            Case "AllFilesystemObjects"
                tsbAllObjects.Checked = True
        End Select
    End Sub
End Class