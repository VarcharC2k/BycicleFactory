namespace VTMES3_RE.View.Systems
{
    partial class frmAuthorGroup
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAuthorGroup));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.MenuTreeList = new DevExpress.XtraTreeList.TreeList();
            this.colCorpId = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.colTitle_ko = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.colRowNum = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.colGroupLevel = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.colId = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.colParentId = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.colGroupYn = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.colProjectName = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.colChkAuthor1 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.chkAuthor1 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
            this.colChkAuthor2 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.chkAuthor2 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
            this.colCode1 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.colCode2 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.menuGroupBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.codeDataSet = new VTMES3_RE.CodeDataSet();
            this.imageCollection11 = new DevExpress.Utils.ImageCollection(this.components);
            this.tileNavPane = new DevExpress.XtraBars.Navigation.TileNavPane();
            this.navTitle = new DevExpress.XtraBars.Navigation.NavButton();
            this.cmdInsert = new DevExpress.XtraBars.Navigation.NavButton();
            this.cmdSave = new DevExpress.XtraBars.Navigation.NavButton();
            this.cmdDelete = new DevExpress.XtraBars.Navigation.NavButton();
            this.cmdClose = new DevExpress.XtraBars.Navigation.NavButton();
            this.gcCodeAuthorGroup = new DevExpress.XtraGrid.GridControl();
            this.codeAuthorGroupBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.gvCodeAuthorGroup = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colGroupCode = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colGroupName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colCreDt = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colCreId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colModDt = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colModId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem14 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlGroup2 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlGroup3 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.codeAuthorGroupTableAdapter = new VTMES3_RE.CodeDataSetTableAdapters.CodeAuthorGroupTableAdapter();
            this.menuGroupTableAdapter = new VTMES3_RE.CodeDataSetTableAdapters.MenuGroupTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MenuTreeList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkAuthor1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkAuthor2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.menuGroupBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.codeDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.imageCollection11)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tileNavPane)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcCodeAuthorGroup)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.codeAuthorGroupBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvCodeAuthorGroup)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem14)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.MenuTreeList);
            this.layoutControl1.Controls.Add(this.tileNavPane);
            this.layoutControl1.Controls.Add(this.gcCodeAuthorGroup);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(1183, 800);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // MenuTreeList
            // 
            this.MenuTreeList.ColumnPanelRowHeight = 30;
            this.MenuTreeList.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
            this.colCorpId,
            this.colTitle_ko,
            this.colRowNum,
            this.colGroupLevel,
            this.colId,
            this.colParentId,
            this.colGroupYn,
            this.colProjectName,
            this.colChkAuthor1,
            this.colChkAuthor2,
            this.colCode1,
            this.colCode2});
            this.MenuTreeList.Cursor = System.Windows.Forms.Cursors.Default;
            this.MenuTreeList.DataSource = this.menuGroupBindingSource;
            this.MenuTreeList.KeyFieldName = "Id";
            this.MenuTreeList.Location = new System.Drawing.Point(354, 93);
            this.MenuTreeList.MinWidth = 19;
            this.MenuTreeList.Name = "MenuTreeList";
            this.MenuTreeList.OptionsBehavior.PopulateServiceColumns = true;
            this.MenuTreeList.OptionsView.AutoWidth = false;
            this.MenuTreeList.ParentFieldName = "ParentId";
            this.MenuTreeList.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.chkAuthor1,
            this.chkAuthor2});
            this.MenuTreeList.RowHeight = 30;
            this.MenuTreeList.Size = new System.Drawing.Size(808, 685);
            this.MenuTreeList.StateImageList = this.imageCollection11;
            this.MenuTreeList.TabIndex = 180;
            this.MenuTreeList.GetStateImage += new DevExpress.XtraTreeList.GetStateImageEventHandler(this.MenuTreeList_GetStateImage);
            // 
            // colCorpId
            // 
            this.colCorpId.FieldName = "CorpId";
            this.colCorpId.MinWidth = 19;
            this.colCorpId.Name = "colCorpId";
            this.colCorpId.OptionsColumn.AllowEdit = false;
            // 
            // colTitle_ko
            // 
            this.colTitle_ko.Caption = "메뉴항목명";
            this.colTitle_ko.FieldName = "Title_ko";
            this.colTitle_ko.MinWidth = 19;
            this.colTitle_ko.Name = "colTitle_ko";
            this.colTitle_ko.OptionsColumn.AllowEdit = false;
            this.colTitle_ko.Visible = true;
            this.colTitle_ko.VisibleIndex = 0;
            this.colTitle_ko.Width = 400;
            // 
            // colRowNum
            // 
            this.colRowNum.Caption = "순번";
            this.colRowNum.FieldName = "RowNum";
            this.colRowNum.MinWidth = 19;
            this.colRowNum.Name = "colRowNum";
            this.colRowNum.OptionsColumn.AllowEdit = false;
            this.colRowNum.Width = 50;
            // 
            // colGroupLevel
            // 
            this.colGroupLevel.Caption = "레벨";
            this.colGroupLevel.FieldName = "GroupLevel";
            this.colGroupLevel.MinWidth = 19;
            this.colGroupLevel.Name = "colGroupLevel";
            this.colGroupLevel.OptionsColumn.AllowEdit = false;
            this.colGroupLevel.Width = 50;
            // 
            // colId
            // 
            this.colId.FieldName = "Id";
            this.colId.MinWidth = 19;
            this.colId.Name = "colId";
            this.colId.OptionsColumn.AllowEdit = false;
            this.colId.Width = 90;
            // 
            // colParentId
            // 
            this.colParentId.FieldName = "ParentId";
            this.colParentId.MinWidth = 19;
            this.colParentId.Name = "colParentId";
            this.colParentId.OptionsColumn.AllowEdit = false;
            this.colParentId.Width = 90;
            // 
            // colGroupYn
            // 
            this.colGroupYn.Caption = "그룹여부";
            this.colGroupYn.FieldName = "GroupYn";
            this.colGroupYn.MinWidth = 19;
            this.colGroupYn.Name = "colGroupYn";
            this.colGroupYn.OptionsColumn.AllowEdit = false;
            this.colGroupYn.Width = 59;
            // 
            // colProjectName
            // 
            this.colProjectName.Caption = "프로젝트명";
            this.colProjectName.FieldName = "ProjectName";
            this.colProjectName.MinWidth = 19;
            this.colProjectName.Name = "colProjectName";
            this.colProjectName.OptionsColumn.AllowEdit = false;
            this.colProjectName.Width = 120;
            // 
            // colChkAuthor1
            // 
            this.colChkAuthor1.AppearanceCell.Options.UseTextOptions = true;
            this.colChkAuthor1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colChkAuthor1.AppearanceHeader.Options.UseTextOptions = true;
            this.colChkAuthor1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colChkAuthor1.Caption = "조회";
            this.colChkAuthor1.ColumnEdit = this.chkAuthor1;
            this.colChkAuthor1.FieldName = "Author1";
            this.colChkAuthor1.Name = "colChkAuthor1";
            this.colChkAuthor1.Visible = true;
            this.colChkAuthor1.VisibleIndex = 1;
            this.colChkAuthor1.Width = 85;
            // 
            // chkAuthor1
            // 
            this.chkAuthor1.AutoHeight = false;
            this.chkAuthor1.Name = "chkAuthor1";
            this.chkAuthor1.ValueChecked = "1";
            this.chkAuthor1.ValueGrayed = "3";
            this.chkAuthor1.ValueUnchecked = "0";
            this.chkAuthor1.CheckedChanged += new System.EventHandler(this.chkAuthor1_CheckedChanged);
            // 
            // colChkAuthor2
            // 
            this.colChkAuthor2.AppearanceCell.Options.UseTextOptions = true;
            this.colChkAuthor2.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colChkAuthor2.AppearanceHeader.Options.UseTextOptions = true;
            this.colChkAuthor2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colChkAuthor2.Caption = "편집";
            this.colChkAuthor2.ColumnEdit = this.chkAuthor2;
            this.colChkAuthor2.FieldName = "Author2";
            this.colChkAuthor2.Name = "colChkAuthor2";
            this.colChkAuthor2.Visible = true;
            this.colChkAuthor2.VisibleIndex = 2;
            this.colChkAuthor2.Width = 85;
            // 
            // chkAuthor2
            // 
            this.chkAuthor2.AutoHeight = false;
            this.chkAuthor2.Name = "chkAuthor2";
            this.chkAuthor2.ValueChecked = "1";
            this.chkAuthor2.ValueGrayed = "3";
            this.chkAuthor2.ValueUnchecked = "0";
            this.chkAuthor2.CheckedChanged += new System.EventHandler(this.chkAuthor2_CheckedChanged);
            // 
            // colCode1
            // 
            this.colCode1.Caption = "treeListColumn2";
            this.colCode1.FieldName = "Code1";
            this.colCode1.Name = "colCode1";
            // 
            // colCode2
            // 
            this.colCode2.Caption = "treeListColumn2";
            this.colCode2.FieldName = "Code2";
            this.colCode2.Name = "colCode2";
            // 
            // menuGroupBindingSource
            // 
            this.menuGroupBindingSource.DataMember = "MenuGroup";
            this.menuGroupBindingSource.DataSource = this.codeDataSet;
            // 
            // codeDataSet
            // 
            this.codeDataSet.DataSetName = "CodeDataSet";
            this.codeDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // imageCollection11
            // 
            this.imageCollection11.ImageStream = ((DevExpress.Utils.ImageCollectionStreamer)(resources.GetObject("imageCollection11.ImageStream")));
            this.imageCollection11.Images.SetKeyName(0, "copy.png");
            this.imageCollection11.Images.SetKeyName(1, "folder_close.png");
            this.imageCollection11.Images.SetKeyName(2, "folder_open.png");
            this.imageCollection11.Images.SetKeyName(3, "folder.png");
            this.imageCollection11.Images.SetKeyName(4, "writing.png");
            // 
            // tileNavPane
            // 
            this.tileNavPane.AllowGlyphSkinning = true;
            this.tileNavPane.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
            this.tileNavPane.Appearance.Options.UseFont = true;
            this.tileNavPane.AppearanceHovered.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.999999F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.tileNavPane.AppearanceHovered.Options.UseFont = true;
            this.tileNavPane.AppearanceSelected.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.999999F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.tileNavPane.AppearanceSelected.Options.UseFont = true;
            this.tileNavPane.Buttons.Add(this.navTitle);
            this.tileNavPane.Buttons.Add(this.cmdInsert);
            this.tileNavPane.Buttons.Add(this.cmdSave);
            this.tileNavPane.Buttons.Add(this.cmdDelete);
            this.tileNavPane.Buttons.Add(this.cmdClose);
            // 
            // tileNavCategory1
            // 
            this.tileNavPane.DefaultCategory.Name = "tileNavCategory1";
            // 
            // 
            // 
            this.tileNavPane.DefaultCategory.Tile.DropDownOptions.BeakColor = System.Drawing.Color.Empty;
            this.tileNavPane.Location = new System.Drawing.Point(12, 12);
            this.tileNavPane.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.tileNavPane.Name = "tileNavPane";
            this.tileNavPane.Size = new System.Drawing.Size(1159, 46);
            this.tileNavPane.TabIndex = 12;
            this.tileNavPane.Text = "tileNavPane";
            // 
            // navTitle
            // 
            this.navTitle.Appearance.Font = new System.Drawing.Font("맑은 고딕", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.navTitle.Appearance.Options.UseFont = true;
            this.navTitle.AppearanceHovered.Font = new System.Drawing.Font("맑은 고딕", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.navTitle.AppearanceHovered.Options.UseFont = true;
            this.navTitle.AppearanceSelected.Font = new System.Drawing.Font("맑은 고딕", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.navTitle.AppearanceSelected.Options.UseFont = true;
            this.navTitle.Caption = "권한그룹 관리";
            this.navTitle.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("navTitle.ImageOptions.SvgImage")));
            this.navTitle.Name = "navTitle";
            // 
            // cmdInsert
            // 
            this.cmdInsert.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.cmdInsert.Caption = "그룹등록";
            this.cmdInsert.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("cmdInsert.ImageOptions.Image")));
            this.cmdInsert.Name = "cmdInsert";
            this.cmdInsert.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.cmdInsert_ElementClick);
            // 
            // cmdSave
            // 
            this.cmdSave.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.cmdSave.Caption = "저 장";
            this.cmdSave.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("cmdSave.ImageOptions.Image")));
            this.cmdSave.Name = "cmdSave";
            this.cmdSave.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.cmdSave_ElementClick);
            // 
            // cmdDelete
            // 
            this.cmdDelete.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.cmdDelete.Caption = "삭 제";
            this.cmdDelete.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("cmdDelete.ImageOptions.Image")));
            this.cmdDelete.Name = "cmdDelete";
            this.cmdDelete.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.cmdDelete_ElementClick);
            // 
            // cmdClose
            // 
            this.cmdClose.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.cmdClose.Caption = "닫 기";
            this.cmdClose.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("cmdClose.ImageOptions.Image")));
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.cmdClose_ElementClick);
            // 
            // gcCodeAuthorGroup
            // 
            this.gcCodeAuthorGroup.DataSource = this.codeAuthorGroupBindingSource;
            this.gcCodeAuthorGroup.EmbeddedNavigator.Margin = new System.Windows.Forms.Padding(3, 7, 3, 7);
            this.gcCodeAuthorGroup.Location = new System.Drawing.Point(21, 93);
            this.gcCodeAuthorGroup.MainView = this.gvCodeAuthorGroup;
            this.gcCodeAuthorGroup.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.gcCodeAuthorGroup.Name = "gcCodeAuthorGroup";
            this.gcCodeAuthorGroup.Size = new System.Drawing.Size(311, 685);
            this.gcCodeAuthorGroup.TabIndex = 13;
            this.gcCodeAuthorGroup.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvCodeAuthorGroup});
            // 
            // codeAuthorGroupBindingSource
            // 
            this.codeAuthorGroupBindingSource.DataMember = "CodeAuthorGroup";
            this.codeAuthorGroupBindingSource.DataSource = this.codeDataSet;
            // 
            // gvCodeAuthorGroup
            // 
            this.gvCodeAuthorGroup.ColumnPanelRowHeight = 32;
            this.gvCodeAuthorGroup.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colGroupCode,
            this.colGroupName,
            this.colCreDt,
            this.colCreId,
            this.colModDt,
            this.colModId});
            this.gvCodeAuthorGroup.DetailHeight = 374;
            this.gvCodeAuthorGroup.GridControl = this.gcCodeAuthorGroup;
            this.gvCodeAuthorGroup.Name = "gvCodeAuthorGroup";
            this.gvCodeAuthorGroup.NewItemRowText = "클릭하여 권한그룹을 등록하세요.";
            this.gvCodeAuthorGroup.OptionsDetail.EnableMasterViewMode = false;
            this.gvCodeAuthorGroup.OptionsView.ShowGroupPanel = false;
            this.gvCodeAuthorGroup.RowHeight = 32;
            this.gvCodeAuthorGroup.FocusedRowChanged += new DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventHandler(this.gvCodeAuthorGroup_FocusedRowChanged);
            // 
            // colGroupCode
            // 
            this.colGroupCode.Caption = "그룹코드";
            this.colGroupCode.FieldName = "GroupCode";
            this.colGroupCode.MinWidth = 23;
            this.colGroupCode.Name = "colGroupCode";
            this.colGroupCode.OptionsColumn.AllowEdit = false;
            this.colGroupCode.OptionsColumn.AllowFocus = false;
            this.colGroupCode.OptionsColumn.ReadOnly = true;
            this.colGroupCode.OptionsEditForm.Visible = DevExpress.Utils.DefaultBoolean.True;
            this.colGroupCode.Width = 86;
            // 
            // colGroupName
            // 
            this.colGroupName.Caption = "권한그룹명";
            this.colGroupName.FieldName = "GroupName";
            this.colGroupName.MinWidth = 23;
            this.colGroupName.Name = "colGroupName";
            this.colGroupName.Visible = true;
            this.colGroupName.VisibleIndex = 0;
            this.colGroupName.Width = 86;
            // 
            // colCreDt
            // 
            this.colCreDt.FieldName = "CreDt";
            this.colCreDt.MinWidth = 23;
            this.colCreDt.Name = "colCreDt";
            this.colCreDt.OptionsColumn.AllowEdit = false;
            this.colCreDt.OptionsColumn.AllowFocus = false;
            this.colCreDt.OptionsColumn.ReadOnly = true;
            this.colCreDt.Width = 86;
            // 
            // colCreId
            // 
            this.colCreId.FieldName = "CreId";
            this.colCreId.MinWidth = 23;
            this.colCreId.Name = "colCreId";
            this.colCreId.OptionsColumn.AllowEdit = false;
            this.colCreId.OptionsColumn.AllowFocus = false;
            this.colCreId.OptionsColumn.ReadOnly = true;
            this.colCreId.Width = 86;
            // 
            // colModDt
            // 
            this.colModDt.FieldName = "ModDt";
            this.colModDt.MinWidth = 23;
            this.colModDt.Name = "colModDt";
            this.colModDt.OptionsColumn.AllowEdit = false;
            this.colModDt.OptionsColumn.AllowFocus = false;
            this.colModDt.OptionsColumn.ReadOnly = true;
            this.colModDt.Width = 86;
            // 
            // colModId
            // 
            this.colModId.FieldName = "ModId";
            this.colModId.MinWidth = 23;
            this.colModId.Name = "colModId";
            this.colModId.OptionsColumn.AllowEdit = false;
            this.colModId.OptionsColumn.AllowFocus = false;
            this.colModId.OptionsColumn.ReadOnly = true;
            this.colModId.Width = 86;
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem14,
            this.layoutControlGroup2,
            this.layoutControlGroup3});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(1183, 800);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem14
            // 
            this.layoutControlItem14.Control = this.tileNavPane;
            this.layoutControlItem14.ControlAlignment = System.Drawing.ContentAlignment.TopLeft;
            this.layoutControlItem14.CustomizationFormText = "layoutControlItem1";
            this.layoutControlItem14.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem14.MaxSize = new System.Drawing.Size(0, 50);
            this.layoutControlItem14.MinSize = new System.Drawing.Size(104, 50);
            this.layoutControlItem14.Name = "layoutControlItem14";
            this.layoutControlItem14.OptionsPrint.AppearanceItem.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlItem14.OptionsPrint.AppearanceItem.Options.UseFont = true;
            this.layoutControlItem14.OptionsPrint.AppearanceItemControl.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlItem14.OptionsPrint.AppearanceItemControl.Options.UseFont = true;
            this.layoutControlItem14.OptionsPrint.AppearanceItemText.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlItem14.OptionsPrint.AppearanceItemText.Options.UseFont = true;
            this.layoutControlItem14.Size = new System.Drawing.Size(1163, 50);
            this.layoutControlItem14.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem14.Text = "layoutControlItem1";
            this.layoutControlItem14.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem14.TextVisible = false;
            // 
            // layoutControlGroup2
            // 
            this.layoutControlGroup2.CustomizationFormText = "권한 그룹";
            this.layoutControlGroup2.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem2});
            this.layoutControlGroup2.Location = new System.Drawing.Point(0, 50);
            this.layoutControlGroup2.Name = "layoutControlGroup2";
            this.layoutControlGroup2.OptionsPrint.AppearanceGroupCaption.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlGroup2.OptionsPrint.AppearanceGroupCaption.Options.UseFont = true;
            this.layoutControlGroup2.OptionsPrint.AppearanceItem.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlGroup2.OptionsPrint.AppearanceItem.Options.UseFont = true;
            this.layoutControlGroup2.OptionsPrint.AppearanceItemControl.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlGroup2.OptionsPrint.AppearanceItemControl.Options.UseFont = true;
            this.layoutControlGroup2.OptionsPrint.AppearanceItemText.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlGroup2.OptionsPrint.AppearanceItemText.Options.UseFont = true;
            this.layoutControlGroup2.Padding = new DevExpress.XtraLayout.Utils.Padding(6, 6, 7, 7);
            this.layoutControlGroup2.Size = new System.Drawing.Size(333, 730);
            this.layoutControlGroup2.Text = "권한 그룹";
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.gcCodeAuthorGroup;
            this.layoutControlItem2.ControlAlignment = System.Drawing.ContentAlignment.TopLeft;
            this.layoutControlItem2.CustomizationFormText = "layoutControlItem2";
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem2.MaxSize = new System.Drawing.Size(315, 0);
            this.layoutControlItem2.MinSize = new System.Drawing.Size(315, 25);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.OptionsPrint.AppearanceItem.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlItem2.OptionsPrint.AppearanceItem.Options.UseFont = true;
            this.layoutControlItem2.OptionsPrint.AppearanceItemControl.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlItem2.OptionsPrint.AppearanceItemControl.Options.UseFont = true;
            this.layoutControlItem2.OptionsPrint.AppearanceItemText.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlItem2.OptionsPrint.AppearanceItemText.Options.UseFont = true;
            this.layoutControlItem2.Size = new System.Drawing.Size(315, 689);
            this.layoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // layoutControlGroup3
            // 
            this.layoutControlGroup3.CustomizationFormText = "권한 리스트";
            this.layoutControlGroup3.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1});
            this.layoutControlGroup3.Location = new System.Drawing.Point(333, 50);
            this.layoutControlGroup3.Name = "layoutControlGroup3";
            this.layoutControlGroup3.OptionsPrint.AppearanceGroupCaption.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlGroup3.OptionsPrint.AppearanceGroupCaption.Options.UseFont = true;
            this.layoutControlGroup3.OptionsPrint.AppearanceItem.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlGroup3.OptionsPrint.AppearanceItem.Options.UseFont = true;
            this.layoutControlGroup3.OptionsPrint.AppearanceItemControl.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlGroup3.OptionsPrint.AppearanceItemControl.Options.UseFont = true;
            this.layoutControlGroup3.OptionsPrint.AppearanceItemText.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlGroup3.OptionsPrint.AppearanceItemText.Options.UseFont = true;
            this.layoutControlGroup3.Padding = new DevExpress.XtraLayout.Utils.Padding(6, 6, 7, 7);
            this.layoutControlGroup3.Size = new System.Drawing.Size(830, 730);
            this.layoutControlGroup3.Text = "권한 리스트";
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.MenuTreeList;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(812, 689);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // codeAuthorGroupTableAdapter
            // 
            this.codeAuthorGroupTableAdapter.ClearBeforeFill = true;
            // 
            // menuGroupTableAdapter
            // 
            this.menuGroupTableAdapter.ClearBeforeFill = true;
            // 
            // frmAuthorGroup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1183, 800);
            this.Controls.Add(this.layoutControl1);
            this.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.Name = "frmAuthorGroup";
            this.Text = "권한그룹 관리";
            this.Load += new System.EventHandler(this.frmAuthorGroup_Load);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.MenuTreeList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkAuthor1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkAuthor2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.menuGroupBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.codeDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.imageCollection11)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tileNavPane)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcCodeAuthorGroup)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.codeAuthorGroupBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvCodeAuthorGroup)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem14)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraBars.Navigation.TileNavPane tileNavPane;
        private DevExpress.XtraBars.Navigation.NavButton navTitle;
        private DevExpress.XtraBars.Navigation.NavButton cmdInsert;
        private DevExpress.XtraBars.Navigation.NavButton cmdSave;
        private DevExpress.XtraBars.Navigation.NavButton cmdDelete;
        private DevExpress.XtraBars.Navigation.NavButton cmdClose;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem14;
        private DevExpress.XtraGrid.GridControl gcCodeAuthorGroup;
        private DevExpress.XtraGrid.Views.Grid.GridView gvCodeAuthorGroup;
        private DevExpress.XtraGrid.Columns.GridColumn colGroupCode;
        private DevExpress.XtraGrid.Columns.GridColumn colGroupName;
        private DevExpress.XtraGrid.Columns.GridColumn colCreDt;
        private DevExpress.XtraGrid.Columns.GridColumn colCreId;
        private DevExpress.XtraGrid.Columns.GridColumn colModDt;
        private DevExpress.XtraGrid.Columns.GridColumn colModId;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup3;
        private CodeDataSet codeDataSet;
        private System.Windows.Forms.BindingSource codeAuthorGroupBindingSource;
        private CodeDataSetTableAdapters.CodeAuthorGroupTableAdapter codeAuthorGroupTableAdapter;
        private DevExpress.XtraTreeList.TreeList MenuTreeList;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colCorpId;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colTitle_ko;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colRowNum;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colGroupLevel;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colId;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colParentId;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colGroupYn;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colProjectName;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.Utils.ImageCollection imageCollection11;
        private System.Windows.Forms.BindingSource menuGroupBindingSource;
        private CodeDataSetTableAdapters.MenuGroupTableAdapter menuGroupTableAdapter;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colChkAuthor1;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit chkAuthor1;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colChkAuthor2;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit chkAuthor2;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colCode1;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colCode2;
    }
}