using DevExpress.Utils;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Registrator;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace VTMES3_RE.Common
{
    public partial class ucGridControl : GridControl
    {
        XtraForm parentForm = null;
        protected override void RegisterAvailableViewsCore(InfoCollection infoCollection)
        {
            infoCollection.Clear();

            infoCollection.Add(new ucGridInfoRegistrator());
            infoCollection.Add(new ucBandedGridInfoRegistrator());
            infoCollection.Add(new ucAdvBandedGridInfoRegistrator());
        }

        public ucGridControl()
        {
            Dock = System.Windows.Forms.DockStyle.Fill;
            EmbeddedNavigator.Buttons.Append.Visible = false;
            EmbeddedNavigator.Buttons.CancelEdit.Visible = false;
            EmbeddedNavigator.Buttons.Edit.Visible = false;
            EmbeddedNavigator.Buttons.EndEdit.Visible = false;
            EmbeddedNavigator.Buttons.First.Visible = false;
            EmbeddedNavigator.Buttons.Last.Visible = false;
            EmbeddedNavigator.Buttons.Next.Visible = false;
            EmbeddedNavigator.Buttons.NextPage.Visible = false;
            EmbeddedNavigator.Buttons.Prev.Visible = false;
            EmbeddedNavigator.Buttons.PrevPage.Visible = false;
            EmbeddedNavigator.Buttons.Remove.Visible = false;
            Size = new System.Drawing.Size(543, 400);
            TabIndex = 1;
            UseEmbeddedNavigator = true;
            DataSourceChanged += ucGridContral_DataSourceChanged;
        }

        public ucGridControl(XtraForm _form)
        {
            parentForm = _form;

            Dock = System.Windows.Forms.DockStyle.Fill;
            EmbeddedNavigator.Buttons.Append.Visible = false;
            EmbeddedNavigator.Buttons.CancelEdit.Visible = false;
            EmbeddedNavigator.Buttons.Edit.Visible = false;
            EmbeddedNavigator.Buttons.EndEdit.Visible = false;
            EmbeddedNavigator.Buttons.First.Visible = false;
            EmbeddedNavigator.Buttons.Last.Visible = false;
            EmbeddedNavigator.Buttons.Next.Visible = false;
            EmbeddedNavigator.Buttons.NextPage.Visible = false;
            EmbeddedNavigator.Buttons.Prev.Visible = false;
            EmbeddedNavigator.Buttons.PrevPage.Visible = false;
            EmbeddedNavigator.Buttons.Remove.Visible = false;
            Size = new System.Drawing.Size(543, 400);
            TabIndex = 1;
            UseEmbeddedNavigator = true;
            DataSourceChanged += ucGridContral_DataSourceChanged;
        }

        private string Get_FormName()
        {
            string FormName = "";
            Control control = Parent;
            while (!(control is null))
            {
                if (control.GetType().FullName.IndexOf("Form") > -1)
                {
                    FormName = control.Name;
                    break;
                }
                control = control.Parent;
            }
            return FormName;
        }

        public void Grid_Export()
        {
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Excel (.xlsx)|*.xlsx|Excel (.xls)|*.xls|RichText File (.rtf)|*.rtf|Pdf File (.pdf)|*.pdf|Html File (.html)|*.html";
                saveDialog.FileName = DateTime.Now.ToShortDateString() + "_" + Get_FormName();
                if (saveDialog.ShowDialog() != DialogResult.Cancel)
                {
                    string exportFilePath = saveDialog.FileName.Trim();
                    string fileExtenstion = new FileInfo(exportFilePath).Extension;
                    try
                    {
                        switch (fileExtenstion)
                        {
                            case ".xls":
                                ExportToXls(exportFilePath);
                                break;

                            case ".xlsx":
                                ExportToXlsx(exportFilePath);
                                break;

                            case ".rtf":
                                ExportToRtf(exportFilePath);
                                break;

                            case ".pdf":

                                foreach (BaseView view in Views)
                                {
                                    if (view.GetType() == typeof(GridView) || view.GetType() == typeof(BandedGridView) || view.GetType() == typeof(AdvBandedGridView))
                                    {
                                        ((GridView)view).AppearancePrint.HeaderPanel.Font = new Font("Arial Unicode MS", ((GridView)view).AppearancePrint.HeaderPanel.Font.Size);
                                        ((GridView)view).AppearancePrint.Row.Font = new Font("Arial Unicode MS", ((GridView)view).AppearancePrint.Row.Font.Size);
                                    }
                                }

                                ExportToPdf(exportFilePath);

                                break;

                            case ".html":
                                foreach (BaseView view in Views)
                                {
                                    if (view.GetType() == typeof(GridView) || view.GetType() == typeof(BandedGridView) || view.GetType() == typeof(AdvBandedGridView))
                                    {
                                        ((GridView)view).OptionsPrint.AutoWidth = false;
                                    }
                                }
                                ExportToHtml(exportFilePath);
                                break;

                            case ".mht":
                                ExportToMht(exportFilePath);
                                break;

                            default:
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (File.Exists(exportFilePath))
                    {
                        try
                        {
                            //Try to open the file and let windows decide how to open it.
                            System.Diagnostics.Process.Start(exportFilePath);
                        }
                        catch
                        {
                            string msg = "The file could not be opened." + Environment.NewLine + Environment.NewLine + "Path: " + exportFilePath;
                            MessageBox.Show(msg, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        string msg = "The file could not be saved." + Environment.NewLine + Environment.NewLine + "Path: " + exportFilePath;
                        MessageBox.Show(msg, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        public void ucGridContral_DataSourceChanged(object sender, EventArgs e)
        {
            ((ucGridView)MainView).Layout -= ((ucGridView)MainView).gridView_Layout;
            foreach (GridColumn column in ((ColumnView)MainView).Columns)
            {
                column.OptionsColumn.AllowEdit = false;
                column.OptionsColumn.ReadOnly = true;

                foreach (DataRowView row in WrGlobal.dt_Description)
                {
                    if (row["COLUMN_NAME"].ToString() == column.FieldName)
                    {
                        column.Caption = row["COLUMN_DESCRIPTION"].ToString();
                        break;
                    }
                }
            }
             ((ucGridView)MainView).Layout += ((ucGridView)MainView).gridView_Layout;
        }
    }

    public class ucGridInfoRegistrator : GridInfoRegistrator
    {
        public override string ViewName => "GridView";

        public override BaseView CreateView(GridControl gridControl)
        {
            return new ucGridView(gridControl);
        }
    }

    public class ucGridView : GridView
    {
        public enum Col_Type
        {
            Text,
            Code,
            Numeric,
            Date,
            DateTime,
            Invisible
        }

        public ucGridView()
        {
            FocusRectStyle = DrawFocusRectStyle.RowFullFocus;
            OptionsBehavior.AutoExpandAllGroups = true;
            OptionsBehavior.AutoSelectAllInEditor = false;
            OptionsBehavior.Editable = true;
            OptionsFind.AlwaysVisible = true;
            OptionsSelection.EnableAppearanceFocusedCell = false;
            OptionsSelection.UseIndicatorForSelection = false;
            OptionsView.ColumnAutoWidth = false;
            OptionsView.ShowIndicator = false;
            CustomDrawCell += new RowCellCustomDrawEventHandler(gridView_CustomDrawCell);
        }

        public ucGridView(GridControl gridControl) : base(gridControl)
        {
            FocusRectStyle = DrawFocusRectStyle.RowFullFocus;
            GridControl = gridControl;
            OptionsBehavior.AutoExpandAllGroups = true;
            OptionsBehavior.AutoSelectAllInEditor = false;
            OptionsBehavior.Editable = true;
            OptionsFind.AlwaysVisible = true;
            OptionsSelection.EnableAppearanceFocusedCell = false;
            OptionsSelection.UseIndicatorForSelection = false;
            OptionsView.ColumnAutoWidth = false;
            OptionsView.ShowIndicator = false;
            CustomDrawCell += new RowCellCustomDrawEventHandler(gridView_CustomDrawCell);
        }

        protected override GridColumnCollection CreateColumnCollection()
        {
            return new ucGridColumnCollection(this);
        }

        public void gridView_Layout(object sender, EventArgs e)
        {
            SaveLayout();
        }

        public void SaveLayout()
        {
            if (Application.ProductName == "Microsoft® Visual Studio®" || GridControl.DataSource is null)
            {
                return;
            }
            FileInfo fileInfo = new FileInfo(@".\Setting\" + Get_FormName() + "_" + GridControl.Name + "_" + Name + ".xml");
            if (!fileInfo.Exists)
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(@".\Setting");
                if (!directoryInfo.Exists)
                {
                    directoryInfo.Create();
                }

                FileStream fs = fileInfo.Create();
                fs.Close();
                fs.Dispose();
            }
            SaveLayoutToXml(fileInfo.FullName);
        }

        private void gridView_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {
            if ((GetRowCellValue(e.RowHandle, "DELETE_USER") ?? "").ToString() != "")
            {
                e.Appearance.ForeColor = Color.LightGray;
            }
        }

        private string Get_FormName()
        {
            string FormName = "";
            Control control = GridControl;
            while (!(control is null))
            {
                if (control.GetType().FullName.IndexOf("Form") > -1)
                {
                    FormName = control.Name;
                    break;
                }
                control = control.Parent;
            }
            return FormName;
        }

        public void Set_Column_Type(string fieldName, Col_Type col_type)
        {
            Layout -= gridView_Layout;
            Set_Column_Type(fieldName, col_type, "0");
            Layout += gridView_Layout;
        }

        public void Set_Column_Type(string fieldName, Col_Type col_type, string option)
        {
            GridColumn gridColumn = Columns[fieldName];
            switch (col_type)
            {
                case Col_Type.Text:
                    gridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Near;
                    break;

                case Col_Type.Code:
                    gridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center;
                    break;

                case Col_Type.Numeric:
                    gridColumn.DisplayFormat.FormatType = FormatType.Numeric;
                    gridColumn.DisplayFormat.FormatString = "n" + option;
                    break;

                case Col_Type.Date:
                    gridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center;
                    gridColumn.DisplayFormat.FormatType = FormatType.DateTime;
                    gridColumn.DisplayFormat.FormatString = "yyyy-MM-dd";
                    break;

                case Col_Type.DateTime:
                    gridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center;
                    gridColumn.DisplayFormat.FormatType = FormatType.DateTime;
                    gridColumn.DisplayFormat.FormatString = "yyyy-MM-dd HH:mm:ss";
                    break;

                case Col_Type.Invisible:
                    gridColumn.Visible = false;
                    break;
            }
        }

        public void MemoEdit_Column(string fieldName)
        {
            Layout -= gridView_Layout;
            MemoEdit_Column(fieldName, false);
            Layout += gridView_Layout;
        }

        public void MemoEdit_Column(string fieldName, bool Editable)
        {
            GridColumn gridColumn = Columns[fieldName];
            OptionsView.RowAutoHeight = true;
            RepositoryItemMemoEdit repositoryItemMemoEdit = new RepositoryItemMemoEdit();
            gridColumn.ColumnEdit = repositoryItemMemoEdit;
            if (Editable)
            {
                gridColumn.OptionsColumn.AllowEdit = true;
                gridColumn.OptionsColumn.ReadOnly = false;
                gridColumn.AppearanceCell.TextOptions.WordWrap = WordWrap.Wrap;
            }
        }

        public void CheckBox_Column_Add()
        {
            OptionsSelection.MultiSelect = true;
            OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect;
        }

        public void CheckBox_Column(string fieldName)
        {
            Layout -= gridView_Layout;
            GridColumn gridColumn = Columns[fieldName];
            RepositoryItemCheckEdit repositoryItemCheckEdit = new RepositoryItemCheckEdit
            {
                ValueChecked = true,
                ValueUnchecked = false
            };
            gridColumn.ColumnEdit = repositoryItemCheckEdit;
            gridColumn.OptionsColumn.AllowEdit = true;
            gridColumn.OptionsColumn.ReadOnly = false;
            Layout += gridView_Layout;
        }

        public void Link_Column(string fieldName, string formName)
        {
            Layout -= gridView_Layout;
            GridColumn gridColumn = Columns[fieldName];
            RepositoryItemButtonEdit repositoryItemButtonEdit = new RepositoryItemButtonEdit
            {
                AutoHeight = false
            };
            repositoryItemButtonEdit.Buttons[0].Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Search;
            repositoryItemButtonEdit.Buttons[0].IsLeft = true;
            repositoryItemButtonEdit.Buttons[0].Tag = formName;

            repositoryItemButtonEdit.ButtonClick += new DevExpress.XtraEditors.Controls.ButtonPressedEventHandler(repositoryItemButtonEdit_ButtonClick);

            gridColumn.ColumnEdit = repositoryItemButtonEdit;
            gridColumn.OptionsColumn.AllowEdit = true;
            gridColumn.OptionsColumn.ReadOnly = false;
            gridColumn.AppearanceCell.BackColor = Color.Lavender;
            Layout += gridView_Layout;
        }

        private void repositoryItemButtonEdit_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            //ButtonEdit HyperLinkEdit = (ButtonEdit)sender;
            //Program._Main.Chaild_Load(e.Button.Tag.ToString(), HyperLinkEdit.EditValue.ToString());
        }
    }

    public class ucGridColumnCollection : GridColumnCollection
    {
        public ucGridColumnCollection(ColumnView view) : base(view)
        {
        }

        protected override GridColumn CreateColumn()
        {
            return new ucGridColumn();
        }
    }

    public class ucGridColumn : GridColumn
    {
        public ucGridColumn() : base()
        {
        }
    }

    public class ucBandedGridInfoRegistrator : BandedGridInfoRegistrator
    {
        public override string ViewName => "BandedGridView";

        public override BaseView CreateView(GridControl gridControl)
        {
            return new ucBandedGridView(gridControl);
        }
    }

    public class ucBandedGridView : BandedGridView
    {
        public ucBandedGridView()
        {
        }

        public ucBandedGridView(GridControl gridControl) : base(gridControl)
        {
        }

        protected override GridColumnCollection CreateColumnCollection()
        {
            return new ucBandedGridColumnCollection(this);
        }
    }

    public class ucBandedGridColumnCollection : BandedGridColumnCollection
    {
        public ucBandedGridColumnCollection(ColumnView view) : base(view)
        {
        }

        protected override GridColumn CreateColumn()
        {
            return new ucBandedGridColumn();
        }
    }

    public class ucBandedGridColumn : BandedGridColumn
    {
        public ucBandedGridColumn() : base()
        {
        }
    }

    public class ucAdvBandedGridInfoRegistrator : AdvBandedGridInfoRegistrator
    {
        public override string ViewName => "AdvBandedGridView";

        public override BaseView CreateView(GridControl gridControl)
        {
            return new ucAdvBandedGridView(gridControl);
        }
    }

    public class ucAdvBandedGridView : AdvBandedGridView
    {
        public ucAdvBandedGridView()
        {
        }

        public ucAdvBandedGridView(GridControl gridControl) : base(gridControl)
        {
        }

        protected override GridColumnCollection CreateColumnCollection()
        {
            return new ucBandedGridColumnCollection(this);
        }
    }
}