Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.IO
Imports System.Drawing
Imports System.Data
Imports System.Data.Common
Imports System.Data.OleDb

'※Exploter1を使用する。
'※ListViewExを実装しない。
'※編集/検索系は実装しない。
'※印刷設定も実装しない。

Public Class Explorer1
    'treeview の選択するノードを、プログラムによって変更するかどうかを示します
    Private ChangingSelectedNode As Boolean
    ''' <summary>
    ''' 編集中フラグ
    ''' </summary>
    ''' <remarks></remarks>

    Private m_bDirty As Boolean
    ''' <summary>
    ''' ファイルフィルタ
    ''' </summary>
    ''' <remarks></remarks>
    Private m_sFilter As String

    ''' <summary>
    ''' データベースのテーブル
    ''' </summary>
    Private m_table As DataTable

    ''' <summary>
    ''' データベースの行
    ''' </summary>
    Private m_row As DataRow

    ''' <summary>
    ''' ドキュメント ファイル名
    ''' </summary>
    ''' <remarks></remarks>
    Private m_sPathname As String

    ''' <summary>
    ''' マウスボタンの押下位置
    ''' </summary>
    ''' <remarks></remarks>
    Private m_posPosition As Point

    ''' <summary>
    ''' 行位置
    ''' </summary>
    ''' <remarks></remarks>
    Private m_iItem As Integer

    ''' <summary>
    ''' 列位置
    ''' </summary>
    ''' <remarks></remarks>
    Private m_iSubItem As Integer

    ''' <summary>
    ''' 行位置
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property Row() As Integer
        Get
            Return m_iItem
        End Get
        Set(ByVal value As Integer)
            With StatusStrip.Items(1)
                .Text = String.Format(My.Resources.ID_INDICATOR_ROW, 1 + value)
                .Enabled = Not String.IsNullOrEmpty(m_sPathname)
            End With
            m_iItem = value
        End Set
    End Property

    ''' <summary>
    ''' 列位置
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property Col() As Integer
        Get
            Return m_iSubItem
        End Get
        Set(ByVal value As Integer)
            With StatusStrip.Items(2)
                .Text = String.Format(My.Resources.ID_INDICATOR_COL, 1 + value)
                .Enabled = Not String.IsNullOrEmpty(m_sPathname)
            End With
            m_iSubItem = value
        End Set
    End Property

    ''' <summary>
    ''' フォームをメモリーへ展開しようとした場合のイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Explorer1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        m_bDirty = False
        m_sPathname = String.Empty
        m_sFilter = "Excel 97-2003 ブック (*.xls)|*.xls|すべてのファイル (*.*)|*.*"
        m_iItem = 0
        m_iSubItem = 0
        'UI を設定します
        SetUpListViewColumns()
        LoadTree()

        'イメージリストの設定
        ImageList1.ImageSize = New Size(16, 15)
        ImageList2.ImageSize = New Size(32, 32)
        Dim source() As Icon = New Icon() { _
            SystemIcons.Shield, SystemIcons.Error, SystemIcons.Question, SystemIcons.Exclamation, SystemIcons.Information}
        For Each index As Icon In source
            ImageList1.Images.Add(index)
            ImageList2.Images.Add(index)
        Next

        '標準ツールバーのビューグループにアイコンを設定
        LargeIconsToolStripMenuItem.Image = My.Resources.IDB_VIEW_LARGEICON
        SmallIconsToolStripMenuItem.Image = My.Resources.IDB_VIEW_SMALLICON
        ListToolStripMenuItem.Image = My.Resources.IDB_VIEW_LIST
        DetailsToolStripMenuItem.Image = My.Resources.IDB_VIEW_DETAILS

        'プレビューツールバーにアイコンを設定
        With ToolStrip1.Items
            .Item(0).Image = My.Resources.IDB_PREVIEW_PRINT
            .Item(1).Image = My.Resources.IDB_PREVIEW_PREV
            .Item(2).Image = My.Resources.IDB_PREVIEW_NEXT
            .Item(3).Image = My.Resources.IDB_PREVIEW_ONEPAGE
            .Item(4).Image = My.Resources.IDB_PREVIEW_ZOOMIN
            .Item(5).Image = My.Resources.IDB_PREVIEW_ZOOMOUT
            .Item(6).Image = My.Resources.IDB_PREVIEW_CLOSE
        End With

        'ステータスバーの設定
        With StatusStrip.Items
            With CType(.Item(0), ToolStripStatusLabel)
                .Spring = True
                .TextAlign = ContentAlignment.MiddleLeft
                .ImageAlign = ContentAlignment.MiddleLeft
            End With
            .Add(My.Resources.ID_INDICATOR_ROW)
            .Add(My.Resources.ID_INDICATOR_COL)
            .Add(My.Resources.ID_INDICATOR_CAPS)
            .Add(My.Resources.ID_INDICATOR_NUM)
            .Add(My.Resources.ID_INDICATOR_SCRL)
            .Add(My.Resources.ID_INDICATOR_MODIFY)
            .Add(My.Resources.ID_INDICATOR_DATE)
            Dim value() As String = My.Resources.ID_INDICATOR_TIME.Replace("\n", vbLf).Split(vbLf)
            .Add(value(0))
        End With


        FoldersToolStripButton_Click(sender, e)

        NewToolStripMenuItem_Click(sender, e)

        With Timer1
            .Interval = 125
            .Enabled = True
        End With
        With Timer2
            .Interval = 500
            .Enabled = False
        End With
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadTree()
        ' TODO: treeview に項目を追加するコードを追加します

        Dim tvRoot As TreeNode
        Dim tvNode As TreeNode

        tvRoot = Me.TreeView.Nodes.Add("Root")
        tvNode = tvRoot.Nodes.Add("TreeItem1")
        tvNode = tvRoot.Nodes.Add("TreeItem2")
        tvNode = tvRoot.Nodes.Add("TreeItem3")
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadListView()
        ' TODO: treeview で選択された項目に基づき listview に項目を追加するコードを追加します

        Dim lvItem As ListViewItem
        ListView.Items.Clear()

        lvItem = ListView.Items.Add("ListViewItem1")
        lvItem.SubItems.AddRange(New String() {"Column2", "Column3"})

        lvItem = ListView.Items.Add("ListViewItem2")
        lvItem.SubItems.AddRange(New String() {"Column2", "Column3"})

        lvItem = ListView.Items.Add("ListViewItem3")
        lvItem.SubItems.AddRange(New String() {"Column2", "Column3"})
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetUpListViewColumns()
        ' TODO: listview 列を設定するコードを追加します
        ListView.Columns.Add("Column1")
        ListView.Columns.Add("Column2")
        ListView.Columns.Add("Column3")
        SetView(View.Details)
    End Sub

    ''' <summary>
    ''' 「終了」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click, ToolStripButton7.Click
        'アプリケーションを終了します
        Close()
    End Sub

    ''' <summary>
    ''' 「ツール バー」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolBarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolBarToolStripMenuItem.Click
        'toolstrip の表示状態および関連メニュー項目のチェック状態を切り替えます
        ToolBarToolStripMenuItem.Checked = Not ToolBarToolStripMenuItem.Checked
        ToolStrip.Visible = ToolBarToolStripMenuItem.Checked
    End Sub

    ''' <summary>
    ''' 「ステータス バー」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatusBarToolStripMenuItem.Click
        'statusstrip の表示状態および関連メニュー項目のチェック状態を切り替えます
        StatusBarToolStripMenuItem.Checked = Not StatusBarToolStripMenuItem.Checked
        StatusStrip.Visible = StatusBarToolStripMenuItem.Checked
    End Sub

    ''' <summary>
    ''' フォルダ ペインを表示するかどうかを変更します
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ToggleFoldersVisible()
        '最初に、関連メニュー項目のチェック状態を切り替えます
        FoldersToolStripMenuItem.Checked = Not FoldersToolStripMenuItem.Checked

        '同期するフォルダのツール バー ボタンを変更します
        FoldersToolStripButton.Checked = FoldersToolStripMenuItem.Checked

        ' TreeView を含むパネルを縮小します
        Me.SplitContainer.Panel1Collapsed = Not FoldersToolStripMenuItem.Checked
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub FoldersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FoldersToolStripMenuItem.Click
        ToggleFoldersVisible()
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub FoldersToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FoldersToolStripButton.Click
        ToggleFoldersVisible()
    End Sub

    ''' <summary>
    ''' ビューの表示状態の変更処理
    ''' </summary>
    ''' <param name="View"></param>
    ''' <remarks></remarks>
    Private Sub SetView(ByVal View As System.Windows.Forms.View)
        'どのメニュー項目をチェックするかを設定します
        Dim MenuItemToCheck As ToolStripMenuItem = Nothing
        Select Case View
            Case View.Details
                MenuItemToCheck = DetailsToolStripMenuItem
            Case View.LargeIcon
                MenuItemToCheck = LargeIconsToolStripMenuItem
            Case View.List
                MenuItemToCheck = ListToolStripMenuItem
            Case View.SmallIcon
                MenuItemToCheck = SmallIconsToolStripMenuItem
            Case View.Tile
                MenuItemToCheck = TileToolStripMenuItem
            Case Else
                Debug.Fail("予期しないビュー")
                View = View.Details
                MenuItemToCheck = DetailsToolStripMenuItem
        End Select

        '適切なメニュー項目をチェックし、[表示] メニューのその他の項目をすべて解除します
        For Each MenuItem As ToolStripMenuItem In ListViewToolStripButton.DropDownItems
            If MenuItem Is MenuItemToCheck Then
                MenuItem.Checked = True
            Else
                MenuItem.Checked = False
            End If
        Next

        '最後に、要求されたビューを設定します
        ListView.View = View
    End Sub

    ''' <summary>
    ''' 「一覧」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListToolStripMenuItem.Click
        SetView(View.List)
    End Sub

    ''' <summary>
    ''' 「詳細」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub DetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DetailsToolStripMenuItem.Click
        SetView(View.Details)
    End Sub

    ''' <summary>
    ''' 「大きいアイコン」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub LargeIconsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LargeIconsToolStripMenuItem.Click
        SetView(View.LargeIcon)
    End Sub

    ''' <summary>
    ''' 「小さいアイコン」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SmallIconsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SmallIconsToolStripMenuItem.Click
        SetView(View.SmallIcon)
    End Sub

    ''' <summary>
    ''' 「タイル」表示処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TileToolStripMenuItem.Click
        SetView(View.Tile)
    End Sub

    ''' <summary>
    ''' ドキュメントを「開く」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OpenToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Dim OpenFileDialog As New OpenFileDialog
        With OpenFileDialog
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            .Filter = m_sFilter
            Select Case .ShowDialog(Me)
                Case Windows.Forms.DialogResult.OK
                    If Savemodified() Then
                        m_sPathname = .FileName
                        ' TODO: ファイルを開くコードを追加します
                        m_bDirty = False
                        OnInitialUpdate()
                        MessageBox(My.Resources.AFX_IDS_IDLEMESSAGE, CInt(MessageBoxIcon.Information))
                    End If
                Case Else
                    MessageBox(My.Resources.IDS_AFXBARRES_CANCEL, CInt(MessageBoxIcon.Information))
            End Select
        End With
    End Sub

    ''' <summary>
    ''' ドキュメントを「名前をつけて保存」する処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim SaveFileDialog As New SaveFileDialog
        With SaveFileDialog
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            .Filter = m_sFilter
            .FileName = m_sPathname
            Select Case .ShowDialog(Me)
                Case Windows.Forms.DialogResult.OK
                    m_sPathname = .FileName
                    SaveToolStripMenuItem_Click(sender, e)
                Case Else
                    MessageBox(My.Resources.IDS_AFXBARRES_CANCEL, CInt(MessageBoxIcon.Information))
            End Select
        End With
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TreeView_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView.AfterSelect
        ' TO DO: 現在選択されている treeview のノードに基づき listview の内容を変更するコードを追加します
        LoadListView()
    End Sub

    ''' <summary>
    ''' フォームを閉じようとしている場合のイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Explorer1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Dim result = False
        e.Cancel = Not result
        If ToolStrip.Visible Then
            If m_bDirty Then
                If result Then
                    m_bDirty = False
                End If
            Else
                result = System.Windows.Forms.DialogResult.Yes = MessageBox _
                    (My.Resources.AFX_IDP_ASK_TO_EXIT, _
                    CInt(MessageBoxButtons.YesNo) + _
                    CInt(MessageBoxIcon.Question) + _
                    CInt(MessageBoxDefaultButton.Button2))
            End If
            If result Then
                Timer1.Enabled = False
                Timer2.Enabled = False

                'TODO: 開いている子ウィンドウを閉じたり、
                'DBのクローズが必要であれば、
                'ココで終了処理を行う

                e.Cancel = Not result
            End If
        Else
            ListView.Visible = True
            StatusStrip.Visible = True
            PrintPreviewControl1.Visible = False
            TreeView.Visible = True
            ToolStrip1.Visible = False
            ToolStrip.Visible = True
            MenuStrip.Visible = True
            Explorer1_Resize(sender, Nothing)
        End If
    End Sub

    ''' <summary>
    ''' フォームを大きさを変更する場合の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Explorer1_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        If ToolStrip1.Visible Then
            Dim value As Rectangle = ClientRectangle
            value.Y += ToolStrip1.Height
            value.Height -= ToolStrip.Height
            value.Height -= StatusStrip.Height
            PrintPreviewControl1.Location = value.Location
            PrintPreviewControl1.Size = value.Size
        End If
    End Sub

    ''' <summary>
    ''' インターバルタイマーの処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim index As Date = Now

        UndoToolStripMenuItem.Enabled = False
        RedoToolStripMenuItem.Enabled = False
        CutToolStripMenuItem.Enabled = False
        CopyToolStripMenuItem.Enabled = False
        PasteToolStripMenuItem.Enabled = False
        SelectAllToolStripMenuItem.Enabled = False

        With StatusStrip.Items
            .Item(3).Enabled = Control.IsKeyLocked(Keys.Capital)
            .Item(4).Enabled = Control.IsKeyLocked(Keys.NumLock)
            .Item(5).Enabled = Control.IsKeyLocked(Keys.Scroll)
            .Item(6).Enabled = m_bDirty
            .Item(7).Text = index.ToString(My.Resources.ID_INDICATOR_DATE)
            Dim source() As String = My.Resources.ID_INDICATOR_TIME.Replace("\n", vbLf).Split(vbLf)
            .Item(8).Text = index.ToString(source(IIf(index.Millisecond < 500, 0, 1)))
        End With
    End Sub

    ''' <summary>
    ''' ディレイ・ワンショット・タイマーの処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False
        With ListView
            If .LabelEdit Then
                With .Items.Item(Row)
                    .BeginEdit()
                End With
            End If
        End With
    End Sub

    ''' <summary>
    ''' 「新規作成」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        If Savemodified() Then
            m_bDirty = False
            m_sPathname = String.Empty
            OnInitialUpdate()
            MessageBox(My.Resources.AFX_IDS_IDLEMESSAGE, CInt(MessageBoxIcon.Information))
        End If
    End Sub

    ''' <summary>
    ''' ドキュメントを「保存」する処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        If Not String.IsNullOrEmpty(m_sPathname) Then
            ' TODO: 現在のフォームの内容をファイルに保存するためのコードをここに追加します
            SetTitle()
            m_bDirty = False
            MessageBox(My.Resources.AFX_IDS_IDLEMESSAGE, CInt(MessageBoxIcon.Information))
        Else
            SaveAsToolStripMenuItem_Click(sender, e)
        End If
    End Sub

    '' ■「ページ設定」処理は存在しない。「印刷」と印刷プレビュー処理の位置が逆

    ''' <summary>
    ''' 「印刷」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Select Case PrintDialog1.ShowDialog(Me)
            Case Windows.Forms.DialogResult.OK
                MessageBox(My.Resources.AFX_IDS_IDLEMESSAGE, CInt(MessageBoxIcon.Information))
            Case Else
                MessageBox(My.Resources.IDS_AFXBARRES_CANCEL, CInt(MessageBoxIcon.Information))
        End Select
    End Sub

    ''' <summary>
    ''' 「印刷プレビュー」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub PrintPreviewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        MenuStrip.Visible = False
        ToolStrip.Visible = False
        ToolStrip1.Visible = True
        TreeView.Visible = False
        ListView.Visible = False
        PrintPreviewControl1.Visible = True
        StatusStrip.Visible = True
        Explorer1_Resize(sender, e)
    End Sub

    ''' <summary>
    ''' 「元に戻す」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click

    End Sub

    ''' <summary>
    ''' 「やり直す」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub RedoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RedoToolStripMenuItem.Click

    End Sub

    ''' <summary>
    ''' 「切り取り」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click

    End Sub

    ''' <summary>
    ''' 「コピー」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click

    End Sub

    ''' <summary>
    ''' 「貼り付け」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click

    End Sub

    ''' <summary>
    ''' 「すべて選択」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click

    End Sub

    ''' <summary>
    ''' 「オプション」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsToolStripMenuItem.Click

    End Sub

    '' ■「検索」「次へ」「置換」「挿入」「編集」「削除」処理は存在しない

    ''' <summary>
    ''' 「検索」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ContentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContentsToolStripMenuItem.Click
        System.Windows.Forms.Help.ShowHelp(Me, My.Resources.IDS_HELP_FILENAME, String.Empty)
    End Sub

    ''' <summary>
    ''' 「目次」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub IndexToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IndexToolStripMenuItem.Click
        System.Windows.Forms.Help.ShowHelpIndex(Me, My.Resources.IDS_HELP_FILENAME)
    End Sub

    ''' <summary>
    ''' 「ヘルプ検索」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SearchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SearchToolStripMenuItem.Click
        System.Windows.Forms.Help.ShowHelp(Me, My.Resources.IDS_HELP_FILENAME)
    End Sub

    ''' <summary>
    ''' 「索引」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Dim value As AboutBox1 = New AboutBox1()
        If Not value Is Nothing Then
            value.ShowDialog()
        End If
    End Sub

    ''' <summary>
    ''' 印刷プレビューの「印刷」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Close()
        PrintToolStripMenuItem_Click(sender, e)
    End Sub

    ''' <summary>
    ''' 印刷プレビューの「前へ」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click

    End Sub

    ''' <summary>
    ''' 印刷プレビューの「次へ」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click

    End Sub

    ''' <summary>
    ''' 印刷プレビューの「1ページ/2ページ」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click

    End Sub

    ''' <summary>
    ''' 印刷プレビューの「拡大」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click

    End Sub

    ''' <summary>
    ''' 印刷プレビューの「縮小」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click

    End Sub

    ''' <summary>
    ''' 印刷処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

    End Sub

    ''' <summary>
    ''' リストビューの「マウスボタン押下」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListView_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView.MouseDown
        m_posPosition = e.Location
    End Sub

    ''' <summary>
    ''' リストビューの「クリック」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView.Click
        Dim info As ListViewHitTestInfo = ListView.HitTest(m_posPosition)
        If Not info Is Nothing Then
            Row = info.Item.Index
            Col = info.Item.SubItems.IndexOf(info.SubItem)
        End If
    End Sub

    ''' <summary>
    ''' リストビューの「アイテム選択変更」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListView_ItemSelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles ListView.ItemSelectionChanged

    End Sub

    ''' <summary>
    ''' リストビューの「キーボードのキーの押下」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListView_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView.KeyDown

    End Sub

    ''' <summary>
    ''' リストビューの「カラムヘッダの描画」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListView_DrawColumnHeader(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawListViewColumnHeaderEventArgs) Handles ListView.DrawColumnHeader

    End Sub

    ''' <summary>
    ''' リストビューの「アイテム描画」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListView_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawListViewItemEventArgs) Handles ListView.DrawItem

    End Sub

    ''' <summary>
    ''' リストビューの「サブアイテム描画」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListView_DrawSubItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawListViewSubItemEventArgs) Handles ListView.DrawSubItem

    End Sub

    ''' <summary>
    ''' リストビューの「カラムクリック」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListView_ColumnClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView.ColumnClick

    End Sub

    ''' <summary>
    ''' リストビューの「ダブルクリック」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListView_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView.DoubleClick
        With ListView
            If .LabelEdit Then
                With .Items.Item(Row)
                    .BeginEdit()
                End With
            End If
        End With
    End Sub

    ''' <summary>
    ''' リストビューの「ラベル編集開始」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListView_BeforeLabelEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LabelEditEventArgs) Handles ListView.BeforeLabelEdit
        MessageBox(My.Resources.IDP_AFXBARRES_TEXT_IS_REQUIRED, CInt(MessageBoxIcon.Information))
    End Sub

    ''' <summary>
    ''' リストビューの「ラベル編集開始」処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListView_AfterLabelEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LabelEditEventArgs) Handles ListView.AfterLabelEdit
        e.CancelEdit = True
        If Not e.Label Is Nothing Then
            If Not String.IsNullOrEmpty(e.Label) Then
                m_bDirty = True
                e.CancelEdit = False
                MessageBox(My.Resources.IDS_EDIT_MENU, CInt(MessageBoxIcon.Information))
            Else
                MessageBox(My.Resources.IDP_AFXBARRES_TEXT_IS_REQUIRED)
                Timer2.Enabled = True
            End If
        Else
            MessageBox(My.Resources.IDS_AFXBARRES_CANCEL, CInt(MessageBoxIcon.Information))
        End If
    End Sub

    ''' <summary>
    ''' メッセージボックス表示処理
    ''' </summary>
    ''' <param name="caption"></param>
    ''' <param name="style"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function MessageBox(ByVal caption As String, Optional ByVal style As Integer = CInt(MessageBoxIcon.Exclamation)) As DialogResult
        Dim index As Integer = 7
        index = index And style
        Dim button As MessageBoxButtons = CType(index, MessageBoxButtons)
        index = CInt(MessageBoxDefaultButton.Button2) + CInt(MessageBoxDefaultButton.Button3)
        index = index And style
        Dim defbutton = CType(index, MessageBoxDefaultButton)
        index = CInt(MessageBoxIcon.Exclamation) + CInt(MessageBoxIcon.Information)
        index = index And style
        Dim icon As MessageBoxIcon = CType(index, MessageBoxIcon)
        index = index >> 4
        Dim value As String = String.Empty
        Dim source() As String = My.Resources.IDS_MESSAGE_TITLES.Replace("\n", vbLf).Split(vbLf)
        If index < source.Length Then
            value = source(index)
        End If
        With StatusStrip
            With .Items(0)
                .Text = caption
                .Image = ImageList1.Images.Item(index)
            End With
        End With
        If String.IsNullOrEmpty(value) Then
            value = source(4)
        End If
        Dim result As DialogResult = Windows.Forms.DialogResult.OK
        Select Case (style)
            Case CInt(MessageBoxIcon.Information)
            Case Else
                result = System.Windows.Forms.MessageBox.Show _
                    (Me, caption, value, button, icon, defbutton)
        End Select
        Select Case result
            Case Windows.Forms.DialogResult.Cancel, Windows.Forms.DialogResult.No
                MessageBox(My.Resources.IDS_AFXBARRES_CANCEL, CInt(MessageBoxIcon.Information))
            Case Windows.Forms.DialogResult.Abort
                MessageBox(My.Resources.IDS_AFXBARRES_ABORT, CInt(MessageBoxIcon.Information))
        End Select
        Return result
    End Function


    ''' <summary>
    ''' ドキュメントの編集状態の確認処理
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Savemodified() As Boolean
        Dim result As Boolean = True
        If m_bDirty Then
            result = False
            Dim value As String = My.Resources.AFX_IDS_UNTITLED
            If Not String.IsNullOrEmpty(m_sPathname) Then
                value = Path.GetFileName(m_sPathname)
            End If
            value = My.Resources.AFX_IDP_ASK_TO_SAVE.Replace("%1", value)
            Select Case MessageBox(value, _
                CInt(MessageBoxIcon.Question) + _
                CInt(MessageBoxButtons.YesNoCancel) + _
                CInt(MessageBoxDefaultButton.Button3))
                Case Windows.Forms.DialogResult.Yes
                    SaveToolStripMenuItem_Click(Nothing, Nothing)
                    result = Not m_bDirty
                Case Windows.Forms.DialogResult.No
                    result = True
            End Select
        End If
        Return result
    End Function

    ''' <summary>
    ''' フォームへタイトルを設定する処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetTitle()
        Dim value As String = My.Resources.AFX_IDS_UNTITLED
        If Not String.IsNullOrEmpty(m_sPathname) Then
            value = Path.GetFileName(m_sPathname)
        End If
        value = My.Resources.AFX_IDS_OBJ_TITLE_INPLACE.Replace("%1", My.Application.Info.Title).Replace("%2", value)
        Text = value
    End Sub

    ' TODO: ■メニューアイテムに設定されているショートカット情報を文字列で返す

    ''' <summary>
    ''' ビューの表示状態の変更処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OnInitialUpdate()
        Dim result As Boolean = True

        SetTitle()

        OnUpdate()

        Row = 0
        Col = 0
        With ListView
            result = 0 < .Items.Count
            .Enabled = result
            .Enabled = result
            .GridLines = result
            .BackgroundImage = Nothing
            If result Then
                .BackgroundImage = My.Resources.IDB_VIEW_STRIPED
                With .Items(Row)
                    .Selected = True
                    .Focused = True
                    .EnsureVisible()
                End With
                .Focus()
            End If
        End With

        OpenToolStripMenuItem.Enabled = Not result

        NewToolStripMenuItem.Enabled = result
        SaveToolStripMenuItem.Enabled = result
        SaveAsToolStripMenuItem.Enabled = result
        PrintToolStripMenuItem.Enabled = result
        PrintPreviewToolStripMenuItem.Enabled = result
    End Sub

    ''' <summary>
    ''' コマンドボタンの許可/禁止と現在注目中の行列位置の表示
    ''' </summary>
    Private Sub SetCursor()

    End Sub


    ''' <summary>
    ''' ビューの更新処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OnUpdate()
        Dim result As Boolean = True
        Dim value As String = String.Empty
        With ListView
            .Items.Clear()
            .Columns.Clear()
            If Not String.IsNullOrEmpty(m_sPathname) Then
                Dim factory As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.OleDb")
                Using conn As DbConnection = factory.CreateConnection()
                    Dim builder As DbConnectionStringBuilder = factory.CreateConnectionStringBuilder()
                    Select Case (IntPtr.Size)
                        Case 4
                            builder("Provider") = My.Resources.IDS_CONNECTION_STRING_32
                        Case 8
                            builder("Provider") = My.Resources.IDS_CONNECTION_STRING_64
                    End Select
                    builder("Data Source") = m_sPathname
                    builder("Extended Properties") = My.Resources.IDS_CONNECTION_EXTENDED_PROPERTIES
                    conn.ConnectionString = builder.ToString()
                    conn.Open()
                    Using command As DbCommand = conn.CreateCommand()
                        command.CommandText = My.Resources.IDS_SQL_SELECT
                        m_table = New DataTable()
                        If Not m_table Is Nothing Then
                            Using reader As DbDataReader = command.ExecuteReader()
                                m_table.Load(reader)
                                ' リストビューカラムヘッダの構築
                                For Each col_index As DataColumn In m_table.Columns
                                    .Columns.Add(col_index.ColumnName, 100, HorizontalAlignment.Left)
                                Next
                                ' リストビューアイテムの構築
                                result = True
                                Row = 0
                                For Each row_index As DataRow In m_table.Rows
                                    m_row = row_index
                                    result = UpdateData(False)
                                    If Not result Then
                                        Exit For
                                    End If
                                    Row += 1
                                Next
                                m_row = Nothing
                            End Using
                            m_table = Nothing
                        End If
                    End Using
                    conn.Close()
                End Using
            End If
        End With
    End Sub

    ''' <summary>
    ''' ビューの更新/読み出し処理
    ''' </summary>
    ''' <param name="bSave"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function UpdateData(Optional ByVal bSave As Boolean = True) As Boolean
        Return DoDataExchange(bSave)
    End Function

    ''' <summary>
    ''' データの読み書き処理
    ''' </summary>
    ''' <param name="bSaveAndValidate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DoDataExchange(ByVal bSaveAndValidate As Boolean) As Boolean
        Dim result As Boolean = True
        With ListView
            Dim index As Integer = 0
            For Each col_index As DataColumn In m_table.Columns
                Dim value As String = m_row(col_index.ColumnName).ToString()
                result = DDX_ListViewItemText(bSaveAndValidate, index, value)
                If Not result Then
                    Exit For
                End If
                index += 1
            Next
        End With
        Return result
    End Function

    ''' <summary>
    ''' ビューの更新/読み出し処理
    ''' </summary>
    ''' <param name="bSaveAndValidate">ビューの更新処理/読み出しフラグ。</param>
    ''' <param name="iSubItem">列位置</param>
    ''' <param name="value">【参照】読み書きする文字列</param>
    ''' <returns>読み書き成功の場合、偽以外を返す</returns>
    ''' <remarks></remarks>
    Private Function DDX_ListViewItemText(ByVal bSaveAndValidate As String, ByVal iSubItem As Integer, ByRef value As String) As Boolean
        Dim result As Boolean = False
        If bSaveAndValidate Then
            Dim item As ListViewItem = ListView.Items.Item(Row)
            If Not item Is Nothing Then
                If 0 = iSubItem Then
                    value = item.Text
                    result = True
                Else
                    value = item.SubItems.Item(iSubItem).Text
                    result = True
                End If
            End If
        Else
            If 0 = iSubItem AndAlso ListView.Items.Count <= Row Then
                ListView.Items.Add(value, 4)
                result = True
            Else
                Dim item As ListViewItem = ListView.Items(Row)
                If Not item Is Nothing Then
                    If item.SubItems.Count <= iSubItem Then
                        Dim subitem As ListViewItem.ListViewSubItem = item.SubItems.Add(value)
                        If Not subitem Is Nothing Then
                            result = True
                        End If
                    Else
                        item.SubItems(iSubItem).Text = value
                        result = True
                    End If
                End If
            End If
        End If
        Return result
    End Function

End Class
