using DevExpress.Data;
using DevExpress.Utils;
using DevExpress.XtraCharts;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using JPlatform.Client.Controls6;
using JPlatform.Client.CSIGMESBaseform6;
using JPlatform.Client.JBaseForm6;
using JPlatform.Client.Library6.interFace;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CSI.GMES.KP
{
    public partial class MSKP90020A : CSIGMESBaseform6//JERPBaseForm
    {
        public bool _firstLoad = true, _dateLoad = false;
        public MyCellMergeHelper _Helper = null;
        bool _allow_confirm = false;
        public int _tab = 0;
        public DataTable _dtChartSource = null, _dtSummarySource = null;
        public string _state = "MONTH";

        public MSKP90020A()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            _firstLoad = true;

            base.OnLoad(e);
            NewButton = false;
            AddButton = false;
            DeleteRowButton = false;
            SaveButton = true;
            DeleteButton = false;
            PreviewButton = false;
            PrintButton = false;

            cboDate.EditValue = DateTime.Now.ToString();
            cboMonth.EditValue = DateTime.Now.ToString();
            cboMonthT.EditValue = DateTime.Now.ToString();
            panTop.BackColor = Color.FromArgb(240, 240, 240);
            tabControl.BackColor = Color.FromArgb(240, 240, 240);

            InitCombobox();
            FormatLayout();

            _firstLoad = false;
        }

        public void FormatLayout()
        {
            if (_tab.Equals(0))
            {
                if (_state.Equals("MONTH"))
                {
                    panTop.Height = 50;

                    cboMonth.Visible = true;
                    cboDate.Visible = false;

                    cboMonth.Location = new Point(68, 13);
                    cboDate.Location = new Point(551, 13);

                    lbFactory.Visible = false;
                    lbFactory.Location = new Point(206, 8);
                    lbPlant.Visible = false;
                    btnExport.Visible = true;

                    cboFactory.Visible = false;
                    cboFactory.Location = new Point(271, 8);
                    cboPlant.Visible = false;

                    lbDate.Text = "Month";
                    lbDate.Width = 64;
                    btnExport.Location = new Point(600, 13);
                    lbDate.Location = new Point(3, 13);
                    cboDate.Location = new Point(68, 13);

                    lbGreen.Location = new Point(860, 13);
                    lbYellow.Location = new Point(1022, 13);
                    lbRed.Location = new Point(1169, 13);
                    lbBlack.Location = new Point(1316, 13);

                    lbGreen.Text = "90% <= Rate <= 100%";
                    lbYellow.Text = "80% <= Rate < 90%";
                    lbRed.Text = "70% <= Rate < 80%";
                    lbBlack.Text = "Rate < 70%";

                    lbGroup.Visible = true;
                    cboGroup.Visible = true;

                    lbGroup.Location = new Point(200, 13);
                    cboGroup.Location = new Point(265, 13);

                    btnConfirm.Visible = false;
                    btnConfirm.Location = new Point(888, 8);

                    btnUnconfirm.Visible = true;
                    btnUnconfirm.Location = new Point(725, 13);

                    grpCboType.Visible = true;
                    grpCboType.Location = new Point(410, 10);

                    lbDateT.Visible = false;
                    lbDateT.Location = new Point(1012, 41);
                    cboMonthT.Visible = false;
                    cboMonthT.Location = new Point(1012, 40);
                }
                else if (_state.Equals("YEAR"))
                {
                    panTop.Height = 75;

                    btnConfirm.Visible = false;
                    btnUnconfirm.Visible = false;
                    btnExport.Visible = false;

                    cboMonth.Visible = true;
                    cboDate.Visible = false;

                    cboMonth.Location = new Point(104, 40);
                    cboDate.Location = new Point(551, 8);

                    lbFactory.Visible = true;
                    lbFactory.Location = new Point(39, 8);
                    lbPlant.Visible = true;
                    lbPlant.Location = new Point(265, 8);

                    cboFactory.Visible = true;
                    cboFactory.Location = new Point(104, 8);
                    cboPlant.Visible = true;
                    cboPlant.Location = new Point(320, 8);

                    lbDate.Text = "Month From";
                    lbDate.Width = 100;
                    lbDate.Location = new Point(3, 40);

                    lbGreen.Location = new Point(470, 40);
                    lbYellow.Location = new Point(632, 40);
                    lbRed.Location = new Point(779, 40);
                    lbBlack.Location = new Point(926, 40);

                    lbGreen.Text = "90% <= Rate <= 100%";
                    lbYellow.Text = "80% <= Rate < 90%";
                    lbRed.Text = "70% <= Rate < 80%";
                    lbBlack.Text = "Rate < 70%";

                    lbGroup.Visible = false;
                    cboGroup.Visible = false;

                    grpCboType.Visible = true;
                    grpCboType.Location = new Point(470, 5);

                    lbDateT.Visible = true;
                    lbDateT.Location = new Point(245, 40);
                    cboMonthT.Visible = true;
                    cboMonthT.Location = new Point(320, 40);
                }
            }
            else if (_tab.Equals(1))
            {
                panTop.Height = 55;

                cboMonth.Visible = false;
                cboDate.Visible = true;

                cboDate.Location = new Point(68, 13);
                cboMonth.Location = new Point(551, 13);

                lbFactory.Visible = true;
                lbFactory.Location = new Point(206, 8);
                lbPlant.Visible = true;
                btnExport.Visible = false;

                cboFactory.Visible = true;
                cboFactory.Location = new Point(271, 8);
                cboPlant.Visible = true;

                lbDate.Text = "Date";
                lbDate.Width = 64;
                lbDate.Location = new Point(3, 13);
                lbFactory.Location = new Point(206, 13);
                lbPlant.Location = new Point(410, 13);

                cboDate.Location = new Point(68, 13);
                cboFactory.Location = new Point(271, 13);
                cboPlant.Location = new Point(460, 13);

                lbGreen.Location = new Point(730, 13);
                lbYellow.Location = new Point(892, 13);
                lbRed.Location = new Point(1039, 13);
                lbBlack.Location = new Point(1186, 13);

                lbGreen.Text = "90 <= Result <= 100";
                lbYellow.Text = "80 <= Result < 90";
                lbRed.Text = "70 <= Result < 80";
                lbBlack.Text = "Rate < 70";

                lbGroup.Visible = false;
                cboGroup.Visible = false;

                lbGroup.Location = new Point(687, 40);
                cboGroup.Location = new Point(752, 40);

                btnConfirm.Visible = true;
                btnConfirm.Location = new Point(605, 13);

                btnUnconfirm.Visible = false;
                btnUnconfirm.Location = new Point(1004, 8);

                grpCboType.Visible = false;

                lbDateT.Visible = false;
                lbDateT.Location = new Point(1012, 41);
                cboMonthT.Visible = false;
                cboMonthT.Location = new Point(1012, 40);
            }
        }

        #region [Start Button Event Code By UIBuilder]

        public override void QueryClick()
        {
            try
            {
                pbProgressShow();
                _allow_confirm = false;

                if (_tab.Equals(0))
                {
                    if (_state.Equals("MONTH"))
                    {
                        InitControls(grdSummary);
                        DataTable _dtSource = GetData("Q_SUMMARY");
                        DataTable _dtChart = GetData("Q_SUMMARY_CHART");
                        _dtChartSource = null;
                        _dtSummarySource = null;

                        if (_dtChart != null && _dtChart.Rows.Count > 0)
                        {
                            fn_load_chart(_dtChart);
                            _dtChartSource = _dtChart.Copy();
                        }
                        else
                        {
                            chartData.DataSource = null;
                            while (chartData.Series[0].Points.Count > 0)
                            {
                                chartData.Series[0].Points.Clear();
                            }
                        }

                        if (_dtSource != null && _dtSource.Rows.Count > 0)
                        {
                            _dtSummarySource = _dtSource.Copy();
                            var distinctValues = _dtSource.AsEnumerable()
                                    .Select(row => new
                                    {
                                        GRP_NM = row.Field<string>("GRP_NM"),
                                        LINE_CD = row.Field<string>("LINE_CD"),
                                        LINE_NM = row.Field<string>("LINE_NM"),
                                    })
                                    .Distinct();
                            DataTable _dtHead = LINQResultToDataTable(distinctValues).Select("", "").CopyToDataTable();
                            CreateDetailGrid(grdSummary, gvwSummary, _dtHead);
                            DataTable _dtf = Binding_Data(_dtSource);
                            SetData(grdSummary, _dtf);
                            Formart_Grid_Summary();
                        }
                    }
                    else if (_state.Equals("YEAR"))
                    {
                        InitControls(grdSummary);
                        DataTable _dtSource = GetDataPop("Q_YEARLY");
                        DataTable _dtChart = GetDataPop("Q_YEARLY_CHART");
                        _dtChartSource = null;
                        _dtSummarySource = null;

                        if (_dtChart != null && _dtChart.Rows.Count > 0)
                        {
                            fn_load_chart(_dtChart);
                            _dtChartSource = _dtChart.Copy();
                        }
                        else
                        {
                            chartData.DataSource = null;
                            while (chartData.Series[0].Points.Count > 0)
                            {
                                chartData.Series[0].Points.Clear();
                            }
                        }

                        if (_dtSource != null && _dtSource.Rows.Count > 0)
                        {
                            _dtSummarySource = _dtSource.Copy();
                            var distinctValues = _dtSource.AsEnumerable()
                                    .Select(row => new
                                    {
                                        GRP_NM = row.Field<string>("GRP_NM"),
                                        YMD = row.Field<string>("YMD"),
                                        YMD_CAPTION = row.Field<string>("YMD_CAPTION"),
                                    })
                                    .Distinct();
                            DataTable _dtHead = LINQResultToDataTable(distinctValues).Select("", "").CopyToDataTable();
                            CreateDetailGrid(grdSummary, gvwSummary, _dtHead);
                            DataTable _dtf = Binding_Data(_dtSource);
                            SetData(grdSummary, _dtf);
                            Formart_Grid_Summary();
                        }
                    }
                }
                else if (_tab.Equals(1))
                {
                    InitControls(grdDetail);
                    btnConfirm.Enabled = false;
                    DataTable _dtSource = GetData("Q_DETAIL");
                    DataTable _dtCheck = GetData("Q_ALLOW");

                    /////// Disable Save Button
                    if (_dtCheck != null && _dtCheck.Rows.Count > 0)
                    {
                        _allow_confirm = _dtCheck.Rows[0]["CONFIRM_YN"].ToString().Equals("Y") ? false : true;
                        btnConfirm.Enabled = _dtCheck.Rows[0]["CONFIRM_YN"].ToString().Equals("Y") ? false : true;
                    }

                    if (_dtSource != null && _dtSource.Rows.Count > 0)
                    {
                        for (int iRow = 0; iRow < _dtSource.Rows.Count; iRow++)
                        {
                            if (!string.IsNullOrEmpty(_dtSource.Rows[iRow]["GRP_NAME_VN"].ToString()))
                            {
                                byte[] data = Convert.FromBase64String(_dtSource.Rows[iRow]["GRP_NAME_VN"].ToString());
                                string _txt_viet = Encoding.UTF8.GetString(data);
                                _dtSource.Rows[iRow]["GRP_NAME_VN"] = _txt_viet;
                            }

                            if (!string.IsNullOrEmpty(_dtSource.Rows[iRow]["ITEM_NAME_VN"].ToString()))
                            {
                                byte[] data = Convert.FromBase64String(_dtSource.Rows[iRow]["ITEM_NAME_VN"].ToString());
                                string _txt_viet = Encoding.UTF8.GetString(data);
                                _dtSource.Rows[iRow]["ITEM_NAME_VN"] = _txt_viet;
                            }

                            if (!string.IsNullOrEmpty(_dtSource.Rows[iRow]["REASON"].ToString()))
                            {
                                byte[] data = Convert.FromBase64String(_dtSource.Rows[iRow]["REASON"].ToString());
                                string _txt_viet = Encoding.UTF8.GetString(data);
                                string _txt_new = _txt_viet.Replace("@", "\n");
                                _dtSource.Rows[iRow]["REASON"] = _txt_new;
                            }

                            if (_dtSource.Rows[iRow]["GRP_CD"].ToString().ToUpper().Equals("TOTAL"))
                            {
                                _dtSource.Rows[iRow]["ITEM_NAME_VN"] = "";
                                _dtSource.Rows[iRow]["GRP_NAME_VN"] = "Total";
                            }
                        }

                        CreateDetailGrid(grdDetail, gvwDetail, _dtSource);
                        SetData(grdDetail, _dtSource);
                        Formart_Grid_Detail();
                    }
                    else
                    {
                        grdDetail.DataSource = null;
                        gvwDetail.Columns.Clear();
                        gvwDetail.Bands.Clear();
                    }
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
            finally
            {
                pbProgressHide();
            }
        }

        public DataTable Binding_Data(DataTable dtSource)
        {
            try
            {
                DataTable _dtf = GetDataTable(gvwSummary);
                string _col_nm = "", _distinct_row = "";

                for (int iRow = 0; iRow < dtSource.Rows.Count; iRow++)
                {
                    if (!dtSource.Rows[iRow]["DIV"].ToString().Equals(_distinct_row))
                    {
                        _dtf.Rows.Add();
                        _dtf.Rows[_dtf.Rows.Count - 1]["DIV"] = dtSource.Rows[iRow]["DIV"].ToString();

                        _distinct_row = dtSource.Rows[iRow]["DIV"].ToString();
                    }

                    if (_state.Equals("MONTH"))
                    {
                        _col_nm = dtSource.Rows[iRow]["LINE_CD"].ToString();
                    }
                    else if (_state.Equals("YEAR"))
                    {
                        _col_nm = dtSource.Rows[iRow]["YMD"].ToString();
                    }

                    _dtf.Rows[_dtf.Rows.Count - 1][_col_nm] = dtSource.Rows[iRow]["QTY"].ToString();
                }

                return _dtf;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public void fn_load_chart(DataTable _dtSource)
        {
            try
            {
                chartData.DataSource = _dtSource;
                chartData.AnnotationRepository.Clear();

                while (chartData.Series[0].Points.Count > 0)
                {
                    chartData.Series[0].Points.Clear();
                }

                XYDiagram diagOSD = (XYDiagram)chartData.Diagram;
                diagOSD.AxisY.WholeRange.MaxValue = 100;
                AxisX axisX = diagOSD.AxisX;
                axisX.CustomLabels.Clear();

                ConstantLine constantLine1 = diagOSD.AxisY.ConstantLines[0];
                constantLine1.AxisValueSerializable = _dtSource.Rows[0]["TARGET"].ToString();

                if (_state.Equals("MONTH"))
                {
                    for (int i = 0; i < _dtSource.Rows.Count; i++)
                    {
                        string label = _dtSource.Rows[i]["LINE_NM"].ToString();
                        chartData.Series[0].Points.Add(new SeriesPoint(label, _dtSource.Rows[i]["QTY"].ToString()));
                    }
                }
                else if (_state.Equals("YEAR"))
                {
                    for (int i = 0; i < _dtSource.Rows.Count; i++)
                    {
                        string label = _dtSource.Rows[i]["YMD_CAPTION"].ToString();
                        chartData.Series[0].Points.Add(new SeriesPoint(label, _dtSource.Rows[i]["QTY"].ToString()));
                    }
                }

                for (int iRow = 0; iRow < _dtSource.Rows.Count; iRow++)
                {
                    double _rateQty = Double.Parse(_dtSource.Rows[iRow]["QTY"].ToString().Replace("%", ""));

                    if (_rateQty >= 90)
                    {
                        chartData.Series[0].Points[iRow].Color = Color.Green;
                    }
                    else if (_rateQty >= 80 && _rateQty < 90)
                    {
                        chartData.Series[0].Points[iRow].Color = Color.Gold;
                    }
                    else if (_rateQty >= 70 && _rateQty < 80)
                    {
                        chartData.Series[0].Points[iRow].Color = Color.Red;
                    }
                    else
                    {
                        chartData.Series[0].Points[iRow].Color = Color.Black;
                    }
                }

                //if (!cboPlant.EditValue.ToString().Equals("ALL"))
                //{
                //    for (int i = 0; i < _dtSource.Rows.Count; i++)
                //    {
                //        axisX.CustomLabels.Add(new CustomAxisLabel(name: _dtSource.Rows[i]["LINE_NM"].ToString().Replace("_", "").Trim(), value: _dtSource.Rows[i]["LINE_NM"].ToString())
                //        {
                //            TextColor = Color.FromArgb(255, 50, 50, 50),
                //        });
                //    }
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public string FormatNumber(string value)
        {
            return string.IsNullOrEmpty(value) ? "0" : Math.Round(Double.Parse(value), 1).ToString();
        }

        public string FormatText(string value)
        {
            return string.IsNullOrEmpty(value) ? "" : value;
        }

        public override void SaveClick()
        {
            try
            {
                DialogResult dlr;
                int _result_qty = 0, _max_qty = 0; ;

                DataTable _dtf = BindingData(grdDetail, true, false);
                if (_dtf != null && _dtf.Rows.Count > 0)
                {
                    for (int iRow = 0; iRow < _dtf.Rows.Count; iRow++)
                    {
                        _max_qty = string.IsNullOrEmpty(_dtf.Rows[iRow]["MAX_SCORE"].ToString()) ? 0 : Int32.Parse(_dtf.Rows[iRow]["MAX_SCORE"].ToString());
                        _result_qty = string.IsNullOrEmpty(_dtf.Rows[iRow]["RESULT_SCORE"].ToString()) ? 0 : Int32.Parse(_dtf.Rows[iRow]["RESULT_SCORE"].ToString());

                        if (_result_qty < 0 || _result_qty > _max_qty)
                        {
                            MessageBox.Show("Số điểm thực tế phải trong khoảng từ 0 ~ Điểm tối đa!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }

                dlr = MessageBox.Show("Bạn có muốn Save không?", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dlr == DialogResult.Yes)
                {
                    bool result = SaveData("Q_SAVE");
                    if (result)
                    {
                        MessageBoxW("Save successfully!", IconType.Information);
                        QueryClick();
                    }
                    else
                    {
                        MessageBoxW("Save failed!", IconType.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Formart_Grid_Summary()
        {
            try
            {
                grdSummary.BeginUpdate();

                for (int i = 0; i < gvwSummary.Columns.Count; i++)
                {
                    gvwSummary.Columns[i].OptionsColumn.AllowEdit = false;
                    gvwSummary.Columns[i].OptionsColumn.ReadOnly = true;
                    gvwSummary.Columns[i].OptionsColumn.AllowSort = DefaultBoolean.False;
                    gvwSummary.Columns[i].OptionsColumn.AllowMerge = DefaultBoolean.False;

                    gvwSummary.Columns[i].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gvwSummary.Columns[i].AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gvwSummary.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gvwSummary.Columns[i].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gvwSummary.Columns[i].AppearanceCell.TextOptions.WordWrap = WordWrap.Wrap;
                    gvwSummary.Columns[i].AppearanceCell.Font = new System.Drawing.Font("Calibri", 12, FontStyle.Regular);

                    if (gvwSummary.Columns[i].FieldName.ToString().Equals("DIV"))
                    {
                        gvwSummary.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                    }

                    if (!gvwSummary.Columns[i].FieldName.ToString().Equals("DIV") && cboGroup.EditValue.ToString().Equals("A002"))
                    {
                        gvwSummary.Columns[i].Width = 85;
                    }
                }

                gvwSummary.RowHeight = 35;
                grdSummary.EndUpdate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Formart_Grid_Detail()
        {
            try
            {
                grdDetail.BeginUpdate();

                for (int i = 0; i < gvwDetail.Columns.Count; i++)
                {
                    gvwDetail.Columns[i].OptionsColumn.AllowEdit = false;
                    gvwDetail.Columns[i].OptionsColumn.ReadOnly = true;
                    gvwDetail.Columns[i].OptionsColumn.AllowSort = DefaultBoolean.False;

                    gvwDetail.Columns[i].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gvwDetail.Columns[i].AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gvwDetail.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gvwDetail.Columns[i].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gvwDetail.Columns[i].AppearanceCell.TextOptions.WordWrap = WordWrap.Wrap;
                    gvwDetail.Columns[i].AppearanceCell.Font = new System.Drawing.Font("Calibri", 12, FontStyle.Regular);

                    if (gvwDetail.Columns[i].FieldName.ToString().Equals("GRP_NAME_VN"))
                    {
                        gvwDetail.Columns[i].Width = 130;
                        gvwDetail.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        gvwDetail.Columns[i].ColumnEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Equals("ORD_NO"))
                    {
                        gvwDetail.Columns[i].Width = 70;
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Equals("ITEM_NAME_VN"))
                    {
                        gvwDetail.Columns[i].Width = 550;
                        gvwDetail.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                        gvwDetail.Columns[i].ColumnEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Contains("REASON"))
                    {
                        gvwDetail.Columns[i].Width = 280;
                        gvwDetail.Columns[i].ColumnEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();

                        if (_allow_confirm)
                        {
                            gvwDetail.Columns[i].OptionsColumn.AllowEdit = true;
                            gvwDetail.Columns[i].OptionsColumn.ReadOnly = false;
                        }
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Contains("RESULT_SCORE"))
                    {
                        gvwDetail.Columns[i].Width = 110;

                        if (_allow_confirm)
                        {
                            gvwDetail.Columns[i].OptionsColumn.AllowEdit = true;
                            gvwDetail.Columns[i].OptionsColumn.ReadOnly = false;
                        }
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Contains("MAX_SCORE"))
                    {
                        gvwDetail.Columns[i].Width = 100;
                    }
                }

                gvwDetail.RowHeight = 120;
                grdDetail.EndUpdate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion [Start Button Event Code By UIBuilder] 

        #region [Grid]

        public void CreateDetailGrid(GridControlEx gridControl, BandedGridViewEx gridView, DataTable dtSource)
        {
            //gridControl.Hide();
            gridView.BeginDataUpdate();
            try
            {
                if (_tab.Equals(0))
                {
                    gridControl.DataSource = null;
                    InitControls(gridControl);
                    gridView.Columns.Clear();
                    gridView.Bands.Clear();

                    while (gridView.Columns.Count > 0)
                    {
                        gridView.Columns.RemoveAt(0);
                    }
                    gridView.OptionsView.ShowColumnHeaders = false;

                    GridBandEx gridBand = null;
                    BandedGridColumnEx colBand = new BandedGridColumnEx();

                    gridBand = new GridBandEx() { Caption = dtSource.Rows[0]["GRP_NM"].ToString() };
                    gridView.Bands.Add(gridBand);
                    gridBand.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gridBand.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap;
                    gridBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                    gridBand.AppearanceHeader.Options.UseBackColor = true;

                    colBand = new BandedGridColumnEx() { FieldName = "DIV", Visible = true };
                    colBand.Width = 100;
                    gridBand.Columns.Add(colBand);

                    if (_state.Equals("MONTH"))
                    {
                        for (int iRow = 0; iRow < dtSource.Rows.Count; iRow++)
                        {
                            gridBand = new GridBandEx() { Caption = dtSource.Rows[iRow]["LINE_NM"].ToString() };
                            gridView.Bands.Add(gridBand);
                            gridBand.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap;
                            gridBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                            gridBand.AppearanceHeader.Options.UseBackColor = true;
                            gridBand.RowCount = 2;

                            colBand = new BandedGridColumnEx() { FieldName = dtSource.Rows[iRow]["LINE_CD"].ToString(), Visible = true };
                            colBand.Width = 75;
                            gridBand.Columns.Add(colBand);
                        }
                    }
                    else if (_state.Equals("YEAR"))
                    {
                        for (int iRow = 0; iRow < dtSource.Rows.Count; iRow++)
                        {
                            gridBand = new GridBandEx() { Caption = dtSource.Rows[iRow]["YMD_CAPTION"].ToString() };
                            gridView.Bands.Add(gridBand);
                            gridBand.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap;
                            gridBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                            gridBand.AppearanceHeader.Options.UseBackColor = true;
                            gridBand.RowCount = 2;

                            colBand = new BandedGridColumnEx() { FieldName = dtSource.Rows[iRow]["YMD"].ToString(), Visible = true };
                            colBand.Width = 75;
                            gridBand.Columns.Add(colBand);
                        }
                    }
                }
                else if (_tab.Equals(1))
                {
                    gridControl.DataSource = null;
                    InitControls(gridControl);
                    gridView.Columns.Clear();
                    gridView.Bands.Clear();

                    while (gridView.Columns.Count > 0)
                    {
                        gridView.Columns.RemoveAt(0);
                    }
                    gridView.OptionsView.ShowColumnHeaders = false;

                    GridBandEx gridBand = null;
                    BandedGridColumnEx colBand = new BandedGridColumnEx();
                    int _col_start = Int32.Parse(dtSource.Rows[0]["FIXED_COL_START"].ToString());
                    int _col_end = Int32.Parse(dtSource.Rows[0]["FIXED_COL_END"].ToString());

                    string[] _col_caption = {"Mission\n(Mục tiêu)","Category\n(Hạng mục đánh giá)",
                        "No\n(STT)", "Max Score\n(Chỉ tiêu)", "Actual Score\n(Kết quả)", "Remark\n(Ghi chú lý do)"};
                    string[] _col_field = { "GRP_NAME_VN", "ITEM_NAME_VN", "ORD_NO", "MAX_SCORE", "RESULT_SCORE", "REASON" };

                    for (int iRow = 0; iRow < dtSource.Columns.Count; iRow++)
                    {
                        ////////// KPI Column
                        int iDx = Array.IndexOf(_col_field, dtSource.Columns[iRow].ColumnName.ToString());
                        gridBand = new GridBandEx() { Caption = iDx >= 0 ? _col_caption[iDx] : dtSource.Columns[iRow].ColumnName.ToString() };
                        gridView.Bands.Add(gridBand);

                        if (iRow >= _col_start && iRow <= _col_end)
                        {
                            gridBand.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                        }

                        gridBand.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap;
                        gridBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                        gridBand.AppearanceHeader.Options.UseBackColor = true;
                        gridBand.RowCount = 2;
                        gridBand.Visible = _col_field.Contains(dtSource.Columns[iRow].ColumnName.ToString()) ? true : false;

                        colBand = new BandedGridColumnEx() { FieldName = dtSource.Columns[iRow].ColumnName.ToString(), Visible = _col_field.Contains(dtSource.Columns[iRow].ColumnName.ToString()) };
                        colBand.Width = 120;
                        gridBand.Columns.Add(colBand);
                    }
                }
            }
            catch
            {
                //throw EX;
            }
            //gridControl.Show();
            gridView.EndDataUpdate();
            gridView.ExpandAllGroups();
        }

        private DataTable GetData(string argType)
        {
            try
            {
                P_MSKP90020A_Q proc = new P_MSKP90020A_Q();
                DataTable dtData = null;

                if (argType.Equals("Q_PERMISS"))
                {
                    string _userID = SessionInfo.UserID;
                    dtData = proc.SetParamData(dtData, argType, _userID, "", "");
                }
                else if (argType.Equals("Q_SUMMARY") || argType.Equals("Q_SUMMARY_CHART") || argType.Equals("Q_EXPORT"))
                {
                    string _group = cboGroup.EditValue == null ? "" : cboGroup.EditValue.ToString();
                    dtData = proc.SetParamData(dtData, argType, _group, "", cboMonth.yyyymm);
                }
                else
                {
                    string _factory = cboFactory.EditValue == null ? "" : cboFactory.EditValue.ToString();
                    string _plant = cboPlant.EditValue == null ? "" : cboPlant.EditValue.ToString();

                    dtData = proc.SetParamData(dtData, argType, _factory, _plant, cboDate.yyyymmdd);
                }

                ResultSet rs = CommonCallQuery(dtData, proc.ProcName, proc.GetParamInfo(), false, 90000, "", true);
                if (rs == null || rs.ResultDataSet == null || rs.ResultDataSet.Tables.Count == 0 || rs.ResultDataSet.Tables[0].Rows.Count == 0)
                {
                    return null;
                }
                return rs.ResultDataSet.Tables[0];
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        private DataTable GetDataPop(string argType)
        {
            try
            {
                P_MSKP90020A_POP proc = new P_MSKP90020A_POP();
                DataTable dtData = null;

                string _factory = cboFactory.EditValue == null ? "" : cboFactory.EditValue.ToString();
                string _plant = cboPlant.EditValue == null ? "" : cboPlant.EditValue.ToString();

                dtData = proc.SetParamData(dtData, argType, _factory, _plant, cboMonth.yyyymm, cboMonthT.yyyymm);

                ResultSet rs = CommonCallQuery(dtData, proc.ProcName, proc.GetParamInfo(), false, 90000, "", true);
                if (rs == null || rs.ResultDataSet == null || rs.ResultDataSet.Tables.Count == 0 || rs.ResultDataSet.Tables[0].Rows.Count == 0)
                {
                    return null;
                }
                return rs.ResultDataSet.Tables[0];
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        #endregion [Grid]

        #region [Combobox]

        private void InitCombobox()
        {
            LoadDataCbo(cboFactory, "Factory", "Q_FTY");
            LoadDataCbo(cboPlant, "Plant", "Q_LINE");
            LoadDataCbo(cboGroup, "Group", "Q_GROUP");
        }

        private void LoadDataCbo(LookUpEditEx argCbo, string _cbo_nm, string _type)
        {
            try
            {
                DataTable dt = GetData(_type);
                if (dt == null)
                {
                    argCbo.Properties.Columns.Clear();
                    argCbo.Properties.DataSource = dt;

                    return;
                }

                string columnCode = dt.Columns[0].ColumnName;
                string columnName = dt.Columns[1].ColumnName;
                string captionCode = "Code";
                string captionName = _cbo_nm;

                argCbo.Properties.Columns.Clear();
                argCbo.Properties.DataSource = dt;
                argCbo.Properties.ValueMember = columnCode;
                argCbo.Properties.DisplayMember = columnName;
                argCbo.Properties.Columns.Add(new DevExpress.XtraEditors.Controls.LookUpColumnInfo(columnCode));
                argCbo.Properties.Columns.Add(new DevExpress.XtraEditors.Controls.LookUpColumnInfo(columnName));
                argCbo.Properties.Columns[columnCode].Visible = false;
                argCbo.Properties.Columns[columnCode].Width = 10;
                argCbo.Properties.Columns[columnCode].Caption = captionCode;
                argCbo.Properties.Columns[columnName].Caption = captionName;
                argCbo.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        #endregion [Combobox]

        #region Events

        private void tabControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                int _prev_tab = _tab;
                _tab = tabControl.SelectedTabPageIndex;
                FormatLayout();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cboFactory_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                LoadDataCbo(cboPlant, "Plant", "Q_LINE");
            }
        }

        private void gvwDetail_CellMerge(object sender, CellMergeEventArgs e)
        {
            try
            {
                if (grdDetail.DataSource == null || gvwDetail.RowCount <= 0) return;

                e.Merge = false;
                e.Handled = true;

                if (e.Column.FieldName.ToString().Equals("GRP_NAME_VN"))
                {
                    string _value1 = gvwDetail.GetRowCellValue(e.RowHandle1, e.Column.FieldName.ToString()).ToString();
                    string _value2 = gvwDetail.GetRowCellValue(e.RowHandle2, e.Column.FieldName.ToString()).ToString();

                    if (_value1 == _value2 && !string.IsNullOrEmpty(_value1))
                    {
                        e.Merge = true;
                    }
                }
            }
            catch
            {

            }
        }

        private void gvwDetail_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            try
            {
                if (grdDetail.DataSource == null || gvwDetail.RowCount < 1) return;

                if (e.Column.FieldName.ToString().Contains("RESULT_SCORE"))
                {
                    e.Appearance.BackColor = Color.FromArgb(226, 239, 217);
                }

                if (e.Column.FieldName.ToString().Contains("ORD"))
                {
                    e.Appearance.BackColor = Color.LightYellow;
                }

                if (gvwDetail.GetRowCellValue(e.RowHandle, "GRP_CD").ToString().ToUpper().Equals("TOTAL"))
                {
                    e.Appearance.BackColor = Color.FromArgb(255, 228, 225);
                    e.Appearance.ForeColor = Color.Blue;
                }

                if (gvwDetail.GetRowCellValue(e.RowHandle, "GRP_CD").ToString().ToUpper().Equals("TOTAL"))
                {
                    if (e.Column.FieldName.ToString().Equals("RESULT_SCORE") && !string.IsNullOrEmpty(e.CellValue.ToString()))
                    {
                        double _rateQty = Double.Parse(e.CellValue.ToString());
                        if (_rateQty >= 90)
                        {
                            e.Appearance.BackColor = Color.Green;
                            e.Appearance.ForeColor = Color.White;
                        }
                        else if (_rateQty >= 80 && _rateQty < 90)
                        {
                            e.Appearance.BackColor = Color.Yellow;
                        }
                        else if (_rateQty >= 70 && _rateQty < 80)
                        {
                            e.Appearance.BackColor = Color.Red;
                            e.Appearance.ForeColor = Color.White;
                        }
                        else
                        {
                            e.Appearance.BackColor = Color.Black;
                            e.Appearance.ForeColor = Color.White;
                        }
                    }
                }
            }
            catch { }
        }

        private void gvwDetail_ShowingEditor(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                if (grdDetail.DataSource == null || gvwDetail.RowCount < 1) return;

                if (gvwDetail.FocusedRowHandle == gvwDetail.RowCount - 1)
                {
                    e.Cancel = true;
                }

                if (gvwDetail.GetRowCellValue(gvwDetail.FocusedRowHandle, "ITEM_CD").ToString().ToUpper().Equals("TOTAL"))
                {
                    e.Cancel = true;
                }
            }
            catch { }
        }

        private void gvwDetail_CalcRowHeight(object sender, RowHeightEventArgs e)
        {
            try
            {
                if (grdDetail.DataSource == null || gvwDetail.RowCount < 1) return;

                if (e.RowHandle == gvwDetail.RowCount - 1 || gvwDetail.GetRowCellValue(e.RowHandle, "ITEM_CD").ToString().ToUpper().Equals("TOTAL"))
                {
                    e.RowHeight = 45;
                }

                if (e.RowHandle == 0)
                {
                    e.RowHeight = 190;
                }
            }
            catch { }
        }

        public bool SaveData(string _type)
        {
            try
            {
                bool _result = true;
                DataTable dtData = null;
                P_MSKP90020A_S proc = new P_MSKP90020A_S();
                string machineName = $"{SessionInfo.UserName}|{ Environment.MachineName}|{GetIPAddress()}";
                int iUpdate = 0, iCount = 0;

                pbProgressShow();

                switch (_type)
                {
                    case "Q_SAVE":
                        DataTable _dtf = BindingData(grdDetail, true, false);

                        for (int iRow = 0; iRow < _dtf.Rows.Count; iRow++)
                        {
                            iUpdate++;

                            byte[] dataCounter = System.Text.Encoding.UTF8.GetBytes(_dtf.Rows[iRow]["REASON"].ToString().Trim());
                            string _txt_counter = System.Convert.ToBase64String(dataCounter);

                            dtData = proc.SetParamData(dtData,
                                                 _type,
                                                 _dtf.Rows[iRow]["PLANT_CD"].ToString(),
                                                 _dtf.Rows[iRow]["LINE_CD"].ToString(),
                                                 cboDate.yyyymmdd,
                                                 _dtf.Rows[iRow]["GRP_CD"].ToString(),
                                                 _dtf.Rows[iRow]["ITEM_CD"].ToString(),
                                                 _dtf.Rows[iRow]["MAX_SCORE"].ToString(),
                                                 _dtf.Rows[iRow]["RESULT_SCORE"].ToString(),
                                                 _txt_counter,
                                                 machineName,
                                                 "CSI.GMES.PD.MSPD90229A_S");

                            if (CommonProcessSave(dtData, proc.ProcName, proc.GetParamInfo(), grdDetail))
                            {
                                dtData = null;
                                iCount++;
                            }
                            else
                            {
                                break;
                            }
                        }

                        if (iUpdate == iCount)
                        {
                            _result = true;
                        }
                        else
                        {
                            _result = false;
                        }

                        break;
                    case "Q_CONFIRM":
                        DataTable _dtSource = GetDataTable(gvwDetail);
                        DataTable _dtLink = _dtSource.Select("ORD_NO IS NOT NULL", "").CopyToDataTable();

                        for (int iRow = 0; iRow < _dtLink.Rows.Count; iRow++)
                        {
                            iUpdate++;

                            byte[] dataCounter = System.Text.Encoding.UTF8.GetBytes(_dtLink.Rows[iRow]["REASON"].ToString().Trim());
                            string _txt_counter = System.Convert.ToBase64String(dataCounter);

                            dtData = proc.SetParamData(dtData,
                                                 _type,
                                                 _dtLink.Rows[iRow]["PLANT_CD"].ToString(),
                                                 _dtLink.Rows[iRow]["LINE_CD"].ToString(),
                                                 cboDate.yyyymmdd,
                                                 _dtLink.Rows[iRow]["GRP_CD"].ToString(),
                                                 _dtLink.Rows[iRow]["ITEM_CD"].ToString(),
                                                 _dtLink.Rows[iRow]["MAX_SCORE"].ToString(),
                                                 _dtLink.Rows[iRow]["RESULT_SCORE"].ToString(),
                                                 _txt_counter,
                                                 machineName,
                                                 "CSI.GMES.PD.MSPD90229A_S");

                            if (CommonProcessSave(dtData, proc.ProcName, proc.GetParamInfo(), grdDetail))
                            {
                                dtData = null;
                                iCount++;
                            }
                            else
                            {
                                break;
                            }
                        }

                        if (iUpdate == iCount)
                        {
                            _result = true;
                        }
                        else
                        {
                            _result = false;
                        }
                        break;
                    default:
                        break;
                }

                pbProgressHide();
                return _result;
            }
            catch (Exception ex)
            {
                pbProgressHide();
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        private void radMonthly_CheckedChanged(object sender, EventArgs e)
        {
            if (!_firstLoad && radMonthly.Checked)
            {
                _state = "MONTH";
                FormatLayout();
            }
        }

        private void radYearly_CheckedChanged(object sender, EventArgs e)
        {
            if (!_firstLoad && radYearly.Checked)
            {
                _state = "YEAR";
                FormatLayout();
            }
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            try
            {
                if (grdDetail.DataSource == null || gvwDetail.RowCount < 1) return;

                DialogResult dlr;
                int _result_qty = 0, _max_qty = 0; ;

                DataTable _dtf = GetDataTable(gvwDetail);
                DataTable _dtLink = _dtf.Select("ORD_NO IS NOT NULL", "").CopyToDataTable();

                if (_dtLink != null && _dtLink.Rows.Count > 0)
                {
                    for (int iRow = 0; iRow < _dtLink.Rows.Count; iRow++)
                    {
                        _max_qty = string.IsNullOrEmpty(_dtLink.Rows[iRow]["MAX_SCORE"].ToString()) ? 0 : Int32.Parse(_dtLink.Rows[iRow]["MAX_SCORE"].ToString());
                        _result_qty = string.IsNullOrEmpty(_dtLink.Rows[iRow]["RESULT_SCORE"].ToString()) ? 0 : Int32.Parse(_dtLink.Rows[iRow]["RESULT_SCORE"].ToString());

                        if (_result_qty < 0 || _result_qty > _max_qty)
                        {
                            MessageBox.Show("Số điểm thực tế phải trong khoảng từ 0 ~ Điểm tối đa!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }

                dlr = MessageBox.Show("Bạn có muốn Confirm không?\nLưu ý: Dữ liệu sau khi xác nhận sẽ không được cập nhập nữa!!!", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dlr == DialogResult.Yes)
                {
                    bool result = SaveData("Q_CONFIRM");
                    if (result)
                    {
                        MessageBoxW("Save successfully!", IconType.Information);
                        QueryClick();
                    }
                    else
                    {
                        MessageBoxW("Save failed!", IconType.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void gvwSummary_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            try
            {
                if (grdSummary.DataSource == null || gvwSummary.RowCount < 1) return;

                if (!e.Column.FieldName.ToString().Contains("DIV") && e.RowHandle == gvwSummary.RowCount - 2)
                {
                    if (!string.IsNullOrEmpty(e.CellValue.ToString()))
                    {
                        e.DisplayText = FormatNumber(e.CellValue.ToString()) + "%";
                    }
                }
            }
            catch { }
        }

        private void gvwSummary_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            try
            {
                if (grdSummary.DataSource == null || gvwSummary.RowCount < 1) return;

                if (e.RowHandle == gvwSummary.RowCount - 1 && !e.Column.FieldName.ToString().Contains("DIV"))
                {
                    e.Appearance.BackColor = Color.LightYellow;

                    if (!string.IsNullOrEmpty(e.CellValue.ToString())){
                        double _rateQty = Double.Parse(e.CellValue.ToString());

                        if(_rateQty == 1)
                        {
                            e.Appearance.BackColor = Color.Green;
                            e.Appearance.ForeColor = Color.White;
                        }
                    }
                }

                if(e.RowHandle == gvwSummary.RowCount - 2)
                {
                    if(!e.Column.FieldName.ToString().Equals("DIV") && !string.IsNullOrEmpty(e.CellValue.ToString()))
                    {
                        double _rateQty = Double.Parse(e.CellValue.ToString().Replace("%", ""));

                        if (_rateQty >= 90)
                        {
                            e.Appearance.BackColor = Color.Green;
                            e.Appearance.ForeColor = Color.White;
                        }
                        else if (_rateQty >= 80 && _rateQty < 90)
                        {
                            e.Appearance.BackColor = Color.Yellow;
                        }
                        else if (_rateQty >= 70 && _rateQty < 80)
                        {
                            e.Appearance.BackColor = Color.Red;
                            e.Appearance.ForeColor = Color.White;
                        }
                        else
                        {
                            e.Appearance.BackColor = Color.Black;
                            e.Appearance.ForeColor = Color.White;
                        }
                    }
                }
            }
            catch { }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable _dtPermiss = GetData("Q_PERMISS");
                if (_dtPermiss == null || _dtPermiss.Rows.Count < 1)
                {
                    MessageBox.Show("Bạn không có quyền thực hiện chức năng này!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult dlr = MessageBox.Show("Bạn có muốn Send Email không?", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dlr == DialogResult.Yes)
                {
                    pbProgressShow();

                    P_MSKP90020A_EMAIL proc = new P_MSKP90020A_EMAIL();
                    DataTable dtDataSet = null;
                    dtDataSet = proc.SetParamData(dtDataSet, "Q", cboMonth.yyyymm);

                    ResultSet rs = CommonCallQuery(dtDataSet, proc.ProcName, proc.GetParamInfo(), false, 90000, "", true);
                    if (rs == null || rs.ResultDataSet == null || rs.ResultDataSet.Tables.Count == 0 || rs.ResultDataSet.Tables[0].Rows.Count == 0)
                    {
                        MessageBoxW("Send failed!", IconType.Warning);
                        return;
                    }

                    DataSet dsData = rs.ResultDataSet;
                    DataTable dtData = dsData.Tables[0];
                    DataTable dtChart = dsData.Tables[1];
                    DataTable dtAvg = dsData.Tables[2];
                    DataTable dtHtml = dsData.Tables[3];

                    if (dtData.Rows.Count == 0)
                    {
                        MessageBoxW("Send failed!", IconType.Warning);
                        return;
                    }

                    string _month = dtData.Rows[0]["MM"].ToString();
                    string SubjectName = "PD Final Result";
                    string htmlReturn = GetHtml(dtData, dtHtml, dtAvg);
                    if (htmlReturn == "") return;

                    /////Image List
                    List<string> imgList = new List<string>();
                    string[] _col_field = { "A001", "A002" };

                    for (int iCount = 0; iCount < _col_field.Length; iCount++)
                    {
                        DataTable dtGroup = dtChart.Select("OPTION_CD = '" + _col_field[iCount] + "'", "").CopyToDataTable();
                        string picName = "";

                        if (dtGroup != null && dtGroup.Rows.Count > 0)
                        {
                            bool bChart1 = LoadDataChart(dtGroup);
                            if (!bChart1) return;
                            picName = "LEAN_PD_" + _col_field[iCount];
                            CaptureControl(tlpMain, picName);
                            imgList.Add(picName);
                        }
                    }

                    bool _result = CreateMail(SubjectName, htmlReturn, imgList, "", "", "huynh.it@changshininc.com");
                    pbProgressHide();

                    if (_result)
                    {
                        MessageBoxW("Send successfully!", IconType.Information);
                    }
                    else
                    {
                        MessageBoxW("Send failed!", IconType.Warning);
                    }
                }
            }
            catch
            {
                pbProgressHide();
                throw;
            }
        }

        public bool LoadDataChart(DataTable argDt)
        {
            try
            {
                DataTable _dtSource = argDt.Copy();

                chart1.DataSource = argDt;
                chart1.AnnotationRepository.Clear();

                while (chart1.Series[0].Points.Count > 0)
                {
                    chart1.Series[0].Points.Clear();
                }

                XYDiagram diagOSD = (XYDiagram)chart1.Diagram;
                diagOSD.AxisY.WholeRange.MaxValue = 95;
                AxisX axisX = diagOSD.AxisX;
                axisX.CustomLabels.Clear();

                ConstantLine constantLine1 = diagOSD.AxisY.ConstantLines[0];
                constantLine1.AxisValueSerializable = argDt.Rows[0]["TARGET"].ToString();

                for (int i = 0; i < argDt.Rows.Count; i++)
                {
                    string label = argDt.Rows[i]["LINE_NM"].ToString();
                    chart1.Series[0].Points.Add(new SeriesPoint(label, argDt.Rows[i]["QTY"].ToString()));
                }

                for (int iRow = 0; iRow < argDt.Rows.Count; iRow++)
                {
                    string _rateQty = argDt.Rows[iRow]["STATUS"].ToString();

                    switch (_rateQty)
                    {
                        case "GREEN":
                            chart1.Series[0].Points[iRow].Color = Color.Green;
                            break;
                        case "YELLOW":
                            chart1.Series[0].Points[iRow].Color = Color.Gold;
                            break;
                        case "RED":
                            chart1.Series[0].Points[iRow].Color = Color.Red;
                            break;
                        case "BLACK":
                            chart1.Series[0].Points[iRow].Color = Color.Black;
                            break;
                        default:
                            break;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public void CaptureControl(Control control, string nameImg)
        {
            try
            {
                string Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\GMES_LFA_Capture\";
                Bitmap bmp = new Bitmap(control.Width, control.Height);
                if (!Directory.Exists(Path)) Directory.CreateDirectory(Path);
                control.DrawToBitmap(bmp, new Rectangle(0, 0, control.Width, control.Height));
                bmp.Save(Path + nameImg + @".png", System.Drawing.Imaging.ImageFormat.Png);
            }
            catch
            {

            }
        }

        private bool CreateMail(string Subject, string htmlBody, List<string> imgList, string RecipEmail, string MailCC, string MailBCC)
        {
            try
            {
                bool _result = true;

                Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem mailItem = (Microsoft.Office.Interop.Outlook.MailItem)app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                mailItem.Subject = Subject;
                Microsoft.Office.Interop.Outlook.Recipients oRecips = (Microsoft.Office.Interop.Outlook.Recipients)mailItem.Recipients;

                if (!string.IsNullOrEmpty(RecipEmail))
                {
                    for (int i = 0; i < RecipEmail.Split(';').Length; i++)
                    {
                        Microsoft.Office.Interop.Outlook.Recipient oRecip = (Microsoft.Office.Interop.Outlook.Recipient)oRecips.Add(RecipEmail.Split(';')[i]);
                        oRecip.Resolve();
                    }
                }

                if (MailCC.Trim() != "")
                {
                    mailItem.CC = MailCC;
                }
                if (MailBCC.Trim() != "")
                {
                    mailItem.BCC = MailBCC;
                }

                ////Add Picture
                if (imgList != null)
                {
                    int iPicCount = imgList.Count;
                    string[] imgInfo = new string[iPicCount];
                    StringBuilder strImg = new StringBuilder();
                    string pathPic = "";
                    for (int i = 0; i < iPicCount; i++)
                    {
                        strImg = new StringBuilder();
                        imgInfo[i] = "imgInfo" + (i + 1).ToString();
                        string pic = imgList[i];

                        if (pic.Contains("\\"))
                        {
                            pathPic = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + pic;
                        }
                        else
                        {
                            pic = pic.Contains(".") ? pic : pic + ".png";
                            pathPic = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + $@"\GMES_LFA_Capture\{pic}";
                        }

                        Outlook.Attachment oAttachPic = mailItem.Attachments.Add(pathPic, Outlook.OlAttachmentType.olByValue, null, "tr");
                        oAttachPic.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo[i]);
                        strImg.Append("<br><img src='cid:" + imgInfo[i] + "'>");

                        htmlBody = htmlBody.Replace("{chart" + (i + 1) + "}", strImg.ToString());
                    }
                    mailItem.HTMLBody = htmlBody;
                }
                else
                {
                    mailItem.HTMLBody = htmlBody;
                }

                mailItem.HTMLBody = htmlBody;
                mailItem.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();

                return _result;
            }
            catch 
            {
                return false;
            }
        }

        private string GetHtml(DataTable arg_DtData, DataTable arg_DtHtml, DataTable arg_Avg)
        {
            try
            {
                string htmlReturn = arg_DtHtml.Rows[0]["TEXT1"].ToString();
                htmlReturn = htmlReturn.Replace("{MONTH}", arg_DtData.Rows[0]["YYYYMM"].ToString());
                htmlReturn = htmlReturn.Replace("{MMYYYY}", arg_DtData.Rows[0]["MMYYYY"].ToString());

                string htmlTotal = "";

                for (int iRow = 0; iRow < arg_Avg.Rows.Count; iRow++)
                {
                    string htmlAvg = arg_DtHtml.Rows[0]["TEXT7"].ToString();
                    htmlAvg = htmlAvg.Replace("{GRP_NM}", arg_Avg.Rows[iRow]["GRP_NM"].ToString());
                    htmlAvg = htmlAvg.Replace("{LINE_LIST}", arg_Avg.Rows[iRow]["LINE_LIST"].ToString());
                    htmlAvg = htmlAvg.Replace("{AVG_GROUP}", arg_Avg.Rows[iRow]["TOTAL_AVG"].ToString());
                    htmlAvg = htmlAvg.Replace("{AVG_LINE}", arg_Avg.Rows[iRow]["QTY"].ToString());

                    htmlTotal += htmlAvg;
                }
                htmlReturn = htmlReturn.Replace("{ttotal}", htmlTotal);

                string _distinct_row = "", _columnHtml = "", _distinct_div = "", _row_html = "";
                string _headType = "", _bodyType = "";

                for (int iRow = 0; iRow < arg_DtData.Rows.Count; iRow++)
                {
                    if (!_distinct_row.Equals(arg_DtData.Rows[iRow]["OPTION_CD"].ToString()))
                    {
                        _distinct_row = arg_DtData.Rows[iRow]["OPTION_CD"].ToString();

                        string htmlTableHead = arg_DtHtml.Rows[0]["TEXT2"].ToString();
                        htmlTableHead = htmlTableHead.Replace("{GRP_NM}", arg_DtData.Rows[iRow]["GRP_NM"].ToString());

                        DataTable dtGroup = arg_DtData.Select("OPTION_CD = '" + _distinct_row + "'", "").CopyToDataTable();

                        if (dtGroup != null && dtGroup.Rows.Count > 0)
                        {
                            /////Table Header
                            var distinctValues = dtGroup.AsEnumerable()
                                .Select(row => new
                                {
                                    ORD = row.Field<decimal>("ORD"),
                                    LINE_CD = row.Field<string>("LINE_CD"),
                                    LINE_NM = row.Field<string>("LINE_NM"),
                                })
                                 .Distinct().OrderBy(r => r.ORD);
                            DataTable _dtHead = LINQResultToDataTable(distinctValues).Select("", "").CopyToDataTable();
                            string htmlHead = "";

                            for (int iCount = 0; iCount < _dtHead.Rows.Count; iCount++)
                            {
                                _columnHtml = arg_DtHtml.Rows[0]["TEXT3"].ToString();
                                _columnHtml = _columnHtml.Replace("{LINE_NM}", string.Format("{0}", _dtHead.Rows[iCount]["LINE_NM"].ToString()));
                                htmlHead += _columnHtml;
                            }

                            htmlTableHead = htmlTableHead.Replace("{tPlant}", htmlHead);

                            ///////Table Row
                            string htmlOther = "", htmlCell = "";
                            string htmlTableRow = "";

                            for (int iCount = 0; iCount < dtGroup.Rows.Count; iCount++)
                            {
                                if (!_distinct_div.Equals(dtGroup.Rows[iCount]["DIV"].ToString()))
                                {
                                    _distinct_div = dtGroup.Rows[iCount]["DIV"].ToString();

                                    if (!string.IsNullOrEmpty(htmlOther) && !string.IsNullOrEmpty(_row_html))
                                    {
                                        _row_html = _row_html.Replace("{tRow}", htmlOther);
                                        htmlTableRow += _row_html;
                                    }

                                    htmlOther = "";
                                    _row_html = arg_DtHtml.Rows[0]["TEXT4"].ToString();
                                    _row_html = _row_html.Replace("{DIV}", dtGroup.Rows[iCount]["DIV"].ToString());
                                }

                                if (_distinct_div.ToUpper().Equals("RANK"))
                                {
                                    htmlCell = arg_DtHtml.Rows[0]["TEXT6"].ToString();
                                    htmlCell = htmlCell.Replace("{ALIGN}", "center");
                                }
                                else
                                {
                                    htmlCell = arg_DtHtml.Rows[0]["TEXT5"].ToString();
                                    if (_distinct_div.Contains("%"))
                                    {
                                        htmlCell = htmlCell.Replace("{ALIGN}", "center");
                                    }
                                    else
                                    {
                                        htmlCell = htmlCell.Replace("{ALIGN}", "right");
                                    }
                                }

                                htmlCell = htmlCell.Replace("{QTY}", dtGroup.Rows[iCount]["QTY"].ToString());
                                htmlCell = htmlCell.Replace("{COLOR}", dtGroup.Rows[iCount]["STATUS"].ToString());

                                htmlOther += htmlCell;

                                if (iCount.Equals(dtGroup.Rows.Count - 1))
                                {
                                    _row_html = _row_html.Replace("{tRow}", htmlOther);
                                    htmlTableRow += _row_html;
                                }
                            }

                            ////
                            switch (_distinct_row)
                            {
                                case "A001":
                                    _headType = "thead1";
                                    _bodyType = "tbody1";
                                    break;
                                case "A002":
                                    _headType = "thead2";
                                    _bodyType = "tbody2";
                                    break;
                                default:
                                    break;
                            }
                            htmlReturn = htmlReturn.Replace("{" + _headType + "}", htmlTableHead);
                            htmlReturn = htmlReturn.Replace("{" + _bodyType + "}", htmlTableRow);
                        }
                    }
                }

                return htmlReturn;
            }
            catch
            {
                return "";
            }
        }

        public void FormatNumber_Excel(ExcelRange range)
        {
            foreach (var cell in range)
            {
                string stringValue = cell.Value?.ToString(); // Assuming the cell contains a string value
                double numericValue;
                if (double.TryParse(stringValue, out numericValue))
                {
                    cell.Value = numericValue; // Assign the numeric value to the cell
                }
            }
        }

        private void btnUnconfirm_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable _dtPermiss = GetData("Q_PERMISS");
                if (_dtPermiss == null || _dtPermiss.Rows.Count < 1)
                {
                    MessageBox.Show("Bạn không có quyền thực hiện chức năng này!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                MSKP90020A_POP view = new MSKP90020A_POP();
                view.SetBrowserMain(this._browserMain);
                view.ShowDialog();

                bool _result = view.CheckIsSaved();
                if (_result)
                {
                    QueryClick();
                }
            }
            catch { }
        }

        private void chartData_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                ChartControl chart = sender as ChartControl;
                if (chart == null || chart.DataSource == null) return;

                // Get hit information under the mouse cursor
                ChartHitInfo hitInfo = chart.CalcHitInfo(e.Location);

                // Check if the clicked element is a series point
                if (hitInfo.InSeries && hitInfo.SeriesPoint != null)
                {
                    // Get the argument of the clicked point
                    string _col_nm = hitInfo.SeriesPoint.Argument;

                    for (int iRow = 0; iRow < _dtChartSource.Rows.Count; iRow++)
                    {
                        if (_state.Equals("MONTH"))
                        {
                            if (_dtChartSource.Rows[iRow]["LINE_NM"].ToString() == _col_nm)
                            {
                                string _fty_cd = _dtChartSource.Rows[iRow]["FTY_CD"].ToString();
                                string _line_cd = _dtChartSource.Rows[iRow]["LINE_CD"].ToString();
                                string _date = cboMonth.yyyymm;

                                ////Disable auto change 
                                _firstLoad = true;

                                cboDate.EditValue = _date + "01";
                                cboFactory.EditValue = _fty_cd;
                                LoadDataCbo(cboPlant, "Plant", "Q_LINE");
                                cboPlant.EditValue = _line_cd;
                                tabControl.SelectedTabPageIndex = 1;

                                ////Click event
                                QueryClick();

                                ////Open auto change 
                                _firstLoad = false;

                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void gvwSummary_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            try
            {
                if (grdSummary.DataSource == null || gvwSummary.RowCount < 1) return;

                string _col_nm = e.Column.FieldName.ToString();

                for(int iRow = 0; iRow < _dtSummarySource.Rows.Count; iRow++)
                {
                    if (_state.Equals("MONTH"))
                    {
                        if (_dtSummarySource.Rows[iRow]["LINE_CD"].ToString() == _col_nm)
                        {
                            string _fty_cd = _dtChartSource.Rows[iRow]["FTY_CD"].ToString();
                            string _line_cd = _dtChartSource.Rows[iRow]["LINE_CD"].ToString();
                            string _date = cboMonth.yyyymm;

                            ////Disable auto change 
                            _firstLoad = true;

                            cboDate.EditValue = _date + "01";
                            cboFactory.EditValue = _fty_cd;
                            LoadDataCbo(cboPlant, "Plant", "Q_LINE");
                            cboPlant.EditValue = _line_cd;
                            tabControl.SelectedTabPageIndex = 1;

                            ////Click event
                            QueryClick();

                            ////Open auto change 
                            _firstLoad = false;

                            break;
                        }
                    }
                }

            }
            catch { }
        }

        private void chartData_CustomDrawCrosshair(object sender, CustomDrawCrosshairEventArgs e)
        {
            try
            {
                if (chartData.DataSource == null) return;

                foreach (CrosshairElementGroup group in e.CrosshairElementGroups)
                {
                    CrosshairGroupHeaderElement groupHeaderElement = group.HeaderElement;
                    // Obtain the first series.
                    CrosshairElement element = group.CrosshairElements[0];

                    // Format the text shown for the series in the crosshair cursor label. Specify the text color and marker size. 
                    element.LabelElement.MarkerSize = new Size(15, 15);
                    element.LabelElement.Font = new Font("Tahoma", 8, FontStyle.Bold);
                    element.LabelElement.Text = string.Format("{0}: {1}%", element.SeriesPoint.Argument.Replace("_",""), element.SeriesPoint.Values[0]);
                }
            }
            catch { }
        }

        private void chartData_CustomDrawAxisLabel(object sender, CustomDrawAxisLabelEventArgs e)
        {
            try
            {
                if (chartData.DataSource == null) return;

                if (_state.Equals("YEAR"))
                {
                    AxisBase axis = e.Item.Axis;
                    e.Item.Text = e.Item.Text.Replace("_", "");
                }
            }
            catch { }
        }

        #endregion

        #region Database

        public class P_MSKP90020A_Q : BaseProcClass
        {
            public P_MSKP90020A_Q()
            {
                // Modify Code : Procedure Name
                _ProcName = "LMES.P_MSKP90020A_Q";
                ParamAdd();
            }
            private void ParamAdd()
            {
                _ParamInfo.Add(new ParamInfo("@ARG_WORK_TYPE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_FTY", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_LINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_YMD", "Varchar", 100, "Input", typeof(System.String)));
            }
            public DataTable SetParamData(DataTable dataTable,
                                        System.String ARG_WORK_TYPE,
                                        System.String ARG_FTY,
                                        System.String ARG_LINE,
                                        System.String ARG_YMD)
            {
                if (dataTable == null)
                {
                    dataTable = new DataTable(_ProcName);
                    foreach (ParamInfo pi in _ParamInfo)
                    {
                        dataTable.Columns.Add(pi.ParamName, pi.TypeClass);
                    }
                }
                // Modify Code : Procedure Parameter
                object[] objData = new object[] {
                                                ARG_WORK_TYPE,
                                                ARG_FTY,
                                                ARG_LINE,
                                                ARG_YMD
                };
                dataTable.Rows.Add(objData);
                return dataTable;
            }
        }

        public class P_MSKP90020A_POP : BaseProcClass
        {
            public P_MSKP90020A_POP()
            {
                // Modify Code : Procedure Name
                _ProcName = "LMES.P_MSKP90020A_POP";
                ParamAdd();
            }
            private void ParamAdd()
            {
                _ParamInfo.Add(new ParamInfo("@ARG_WORK_TYPE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_FTY", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_LINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_DATEF", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_DATET", "Varchar", 100, "Input", typeof(System.String)));
            }
            public DataTable SetParamData(DataTable dataTable,
                                        System.String ARG_WORK_TYPE,
                                        System.String ARG_FTY,
                                        System.String ARG_LINE,
                                        System.String ARG_DATEF,
                                        System.String ARG_DATET)
            {
                if (dataTable == null)
                {
                    dataTable = new DataTable(_ProcName);
                    foreach (ParamInfo pi in _ParamInfo)
                    {
                        dataTable.Columns.Add(pi.ParamName, pi.TypeClass);
                    }
                }
                // Modify Code : Procedure Parameter
                object[] objData = new object[] {
                                                ARG_WORK_TYPE,
                                                ARG_FTY,
                                                ARG_LINE,
                                                ARG_DATEF,
                                                ARG_DATET
                };
                dataTable.Rows.Add(objData);
                return dataTable;
            }
        }

        public class P_MSKP90020A_EMAIL : BaseProcClass
        {
            public P_MSKP90020A_EMAIL()
            {
                // Modify Code : Procedure Name
                _ProcName = "LMES.P_MSKP90020A_EMAIL";
                ParamAdd();
            }
            private void ParamAdd()
            {
                _ParamInfo.Add(new ParamInfo("@V_P_SEND_TYPE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@V_P_SEND_DATE", "Varchar", 100, "Input", typeof(System.String)));
            }
            public DataTable SetParamData(DataTable dataTable,
                                        System.String V_P_SEND_TYPE,
                                        System.String V_P_SEND_DATE)
            {
                if (dataTable == null)
                {
                    dataTable = new DataTable(_ProcName);
                    foreach (ParamInfo pi in _ParamInfo)
                    {
                        dataTable.Columns.Add(pi.ParamName, pi.TypeClass);
                    }
                }
                // Modify Code : Procedure Parameter
                object[] objData = new object[] {
                                                V_P_SEND_TYPE,
                                                V_P_SEND_DATE
                };
                dataTable.Rows.Add(objData);
                return dataTable;
            }
        }

        public class P_MSKP90020A_S : BaseProcClass
        {
            public P_MSKP90020A_S()
            {
                // Modify Code : Procedure Name
                _ProcName = "LMES.P_MSKP90020A_S";
                ParamAdd();
            }
            private void ParamAdd()
            {
                _ParamInfo.Add(new ParamInfo("@ARG_TYPE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_PLANT", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_LINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_YMD", "Varchar2", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_GROUP", "Varchar2", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_ITEM", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_MAX_SCORE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_RESULT", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_REASON", "Varchar2", 0, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_CREATE_PC", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_CREATE_PROGRAM_ID", "Varchar", 100, "Input", typeof(System.String)));
            }
            public DataTable SetParamData(DataTable dataTable,
                                        System.String ARG_TYPE,
                                        System.String ARG_PLANT,
                                        System.String ARG_LINE,
                                        System.String ARG_YMD,
                                        System.String ARG_GROUP,
                                        System.String ARG_ITEM,
                                        System.String ARG_MAX_SCORE,
                                        System.String ARG_RESULT,
                                        System.String ARG_REASON,
                                        System.String ARG_CREATE_PC,
                                        System.String ARG_CREATE_PROGRAM_ID)
            {
                if (dataTable == null)
                {
                    dataTable = new DataTable(_ProcName);
                    foreach (ParamInfo pi in _ParamInfo)
                    {
                        dataTable.Columns.Add(pi.ParamName, pi.TypeClass);
                    }
                }
                // Modify Code : Procedure Parameter
                object[] objData = new object[] {
                    ARG_TYPE,
                    ARG_PLANT,
                    ARG_LINE,
                    ARG_YMD,
                    ARG_GROUP,
                    ARG_ITEM,
                    ARG_MAX_SCORE,
                    ARG_RESULT,
                    ARG_REASON,
                    ARG_CREATE_PC,
                    ARG_CREATE_PROGRAM_ID
                };
                dataTable.Rows.Add(objData);
                return dataTable;
            }
        }

        #endregion

        DataTable GetDataTable(GridView view)
        {
            DataTable dt = new DataTable();
            foreach (GridColumn c in view.Columns)
                dt.Columns.Add(c.FieldName, c.ColumnType);
            for (int r = 0; r < view.RowCount; r++)
            {
                object[] rowValues = new object[dt.Columns.Count];
                for (int c = 0; c < dt.Columns.Count; c++)
                    rowValues[c] = view.GetRowCellValue(r, dt.Columns[c].ColumnName);
                dt.Rows.Add(rowValues);
            }
            return dt;
        }

        private DataTable LINQResultToDataTable<T>(IEnumerable<T> Linqlist)
        {
            DataTable dt = new DataTable();
            PropertyInfo[] columns = null;
            if (Linqlist == null) return dt;
            foreach (T Record in Linqlist)
            {
                if (columns == null)
                {
                    columns = ((Type)Record.GetType()).GetProperties();
                    foreach (PropertyInfo GetProperty in columns)
                    {
                        Type colType = GetProperty.PropertyType;

                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition()
                        == typeof(Nullable<>)))
                        {
                            colType = colType.GetGenericArguments()[0];
                        }

                        dt.Columns.Add(new DataColumn(GetProperty.Name, colType));
                    }
                }
                DataRow dr = dt.NewRow();
                foreach (PropertyInfo pinfo in columns)
                {
                    dr[pinfo.Name] = pinfo.GetValue(Record, null) == null ? DBNull.Value : pinfo.GetValue
                    (Record, null);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}