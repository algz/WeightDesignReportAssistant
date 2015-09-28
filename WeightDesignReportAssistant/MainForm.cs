using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Imaging;
using System.Windows.Forms.DataVisualization.Charting;
using System.Collections;
using CatiaCommon;

namespace WeightDesignReportAssistant
{
    public partial class MainForm : Form
    {
        PartProperties[] partProperties = null;

       // private string[] FontSizeName = { "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号" };
        //private string[] FontSizeName = { "8", "9", "10", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72", "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号" };
        
        //private float[] FontSize = { 8, 9, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, 42, 36, 26, 24, 22, 18, 16, 15, 14, 12, 10.5F, 9, 7.5F, 6.5F, 5.5F, 5 };//定义字号数组

        private void MainForm_Load(object sender, EventArgs e)
        {
            //定义字体名称
            //System.Drawing.Text.InstalledFontCollection objFont = new System.Drawing.Text.InstalledFontCollection();
            //foreach (System.Drawing.FontFamily i in objFont.Families)
            //{
            //    this.fontNameComboBox.Items.Add(i.Name.ToString());
            //}
            //string[] fontNames = new string[] { "Arial", "仿宋","华文中宋","华文仿宋","华文新魏","华文楷体","华文琥珀","华文细黑",
            //"华文行楷","华文隶体","宋体","幼圆","微软雅黑","新宋体","方正姚体","方正舒体","楷体","隶书","黑体"};
            //this.fontNameComboBox.Items.AddRange(fontNames);
            //this.fontNameComboBox.SelectedItem = this.docTemplateRichTextBox.Font.Name;

            ////定义字体大小
            //foreach (string name in FontSizeName)
            //{
            //    this.fontSizeComboBox.Items.Add(name);
            //}
            //this.fontSizeComboBox.SelectedItem = this.docTemplateRichTextBox.Font.Size.ToString();
        }

        public MainForm()
        {
            InitializeComponent();

        }

        private string pwdFilePath = "";//项目文件路径

        /// <summary>
        /// 新建项目
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void newPWDMenuItem_Click(object sender, EventArgs e)
        {
            if (this.partProperties != null && MessageBox.Show("是否保存数据", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                this.savePWDMenuItem_Click(null, null);
            }

            this.partProperties = null;
            this.pwdFilePath = "";

            this.partPropertiesGridView.Rows.Clear();
            this.partPropertiesTreeView.Nodes.Clear();
            this.partPropertiesNameText.Text = "";

            this.docTemplateFilePath = "";
            this.docTemplateNameLabel.Text = "";
            this.docTextContextMenu.Items.Clear();
            this.docTemplateTreeView.Nodes.Clear();
            //this.docTemplateRichTextBox.Text = "";
            this.saveDocTemplateBtn.Enabled = false;
            this.exportDocBtn.Enabled = false;

            this.excelTemplateGridView.Rows.Clear();
            this.excelTemplateFilePath = "";
            this.excelTemplateNameLabel.Text = "";
            this.excelTemplateTreeView.Nodes.Clear();
            this.saveExcelTemplateBtn.Enabled = false;
            this.exportExcelBtn.Enabled = false;
            this.customColumnComboBox.Items.Clear();

            this.generatedPicturesTreeView.Nodes.Clear();
            this.picChart.Series[0].Points.Clear();
            this.exportPicBtn.Enabled = false;
        }

        /// <summary>
        /// 打开项目文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openPWDMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = Common.AppPath;
            dialog.Filter = "项目文件(*.pwd)|*.pwd";
            dialog.RestoreDirectory = true;
            dialog.FilterIndex = 1;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.partProperties = Common.ConvertPWDXmlToPartPropertiesArray(dialog.FileName);
                this.pwdFilePath = dialog.FileName;

                if (this.partProperties != null)
                {
                    loadObtainCatia();
                }
                else
                {
                    MessageBox.Show("数据加载不成功");
                }
            }


        }

        /// <summary>
        /// 保存项目文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void savePWDMenuItem_Click(object sender, EventArgs e)
        {
            if (this.partProperties != null && this.pwdFilePath != "")
            {
                Common.savePWDToXml(this.pwdFilePath, this.partProperties);
                MessageBox.Show("项目文件保存成功");
            }
            else
            {
                saveAsPWDMenuItem_Click(null, null);
            }
        }

        private void saveAsPWDMenuItem_Click(object sender, EventArgs e)
        {
            if (this.partProperties != null)
            {
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.InitialDirectory = Common.AppPath;
                dialog.Filter = "文档模板文件(*.pwd)|*.pwd";
                dialog.FileName = "*.pwd";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    Common.savePWDToXml(dialog.FileName, this.partProperties);
                }
            }
            else
            {
                MessageBox.Show("请加载项目文件");
            }
        }

        private void ExitPWDMenuItem_Click(object sender, EventArgs e)
        {
            if (this.partProperties != null && MessageBox.Show("是否保存数据", "提示") == DialogResult.OK)
            {
                savePWDMenuItem_Click(this, null);
            }
            this.Close();
        }

        #region 质量获取

        /// <summary>
        /// 从Catia接口获取数据进行转换
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void obtainCatiaBtn_Click(object sender, EventArgs e)
        {
            try
            {
                CatiaPickGravity CommnonFun = new CatiaPickGravity();
                this.partPropertiesNameText.Text = CommnonFun.GetCatiaModel();
                List<CatiaCommon.PartProperties> lstGravity = CommnonFun.GetGravityInfo();
                if (lstGravity.Count == 0)
                {
                    MessageBox.Show("Catia获取失败.");
                    return;
                }

                this.partProperties = Common.ConvertCatiaInterfaceToPartPropertiesArray(lstGravity);

                if (this.partProperties != null)
                {
                    loadObtainCatia();
                }
                else
                {
                    MessageBox.Show("数据加载不成功");
                }
            }
            catch(Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            

        }
        #endregion

        #region 文档段落生成
        private void partIDCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox box = (CheckBox)sender;
            switch (box.Tag.ToString())
            {
                case "allSelect":
                    foreach (Control c in this.dataPanel.Controls)
                    {
                        CheckBox temBox = c as CheckBox;
                        if (temBox != null)
                        {
                            temBox.Checked = box.Checked;
                        }
                    }
                    break;
                case "partInertiaMatrix":
                    this.partPropertiesGridView.Columns["partInertiaMatrixIXX"].Visible = box.Checked;
                    this.partPropertiesGridView.Columns["partInertiaMatrixIXY"].Visible = box.Checked;
                    this.partPropertiesGridView.Columns["partInertiaMatrixIXZ"].Visible = box.Checked;
                    this.partPropertiesGridView.Columns["partInertiaMatrixIYY"].Visible = box.Checked;
                    this.partPropertiesGridView.Columns["partInertiaMatrixIYX"].Visible = box.Checked;
                    this.partPropertiesGridView.Columns["partInertiaMatrixIYZ"].Visible = box.Checked;
                    this.partPropertiesGridView.Columns["partInertiaMatrixIZX"].Visible = box.Checked;
                    this.partPropertiesGridView.Columns["partInertiaMatrixIZY"].Visible = box.Checked;
                    this.partPropertiesGridView.Columns["partInertiaMatrixIZZ"].Visible = box.Checked;
                    break;
                case "partCenterOfGravity":
                    this.partPropertiesGridView.Columns["partCenterOfGravityX"].Visible = box.Checked;
                    this.partPropertiesGridView.Columns["partCenterOfGravityY"].Visible = box.Checked;
                    this.partPropertiesGridView.Columns["partCenterOfGravityZ"].Visible = box.Checked;
                    break;
                default:
                    this.partPropertiesGridView.Columns[box.Tag.ToString()].Visible = box.Checked;
                    break;
            }
        }

        #region 文档模板Btn
        private string docTemplateFilePath = "";//文档模板文件路径

        /// <summary>
        /// 新建文档模板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void newDocTemplateBtn_Click(object sender, EventArgs e)
        {

            if (this.dxRichTextBox.Text != "")
            {
                if (MessageBox.Show("是否保存模板", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    saveDocTemplateBtn_Click(null, null);
                }
            }
            this.docTemplateFilePath = "";
            this.docTemplateNameLabel.Text = "";
        }

        /// <summary>
        /// 打开文档模板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openDocTemplateBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = Common.DocTemplatePath;
            dialog.Filter = "文档模板文件(*.doc)|*.doc";
            dialog.RestoreDirectory = true;
            dialog.FilterIndex = 1;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.dxRichTextBox.OpenDocument(dialog.FileName);
                this.docTemplateNameLabel.Text = Path.GetFileNameWithoutExtension(dialog.FileName);
                this.saveDocTemplateBtn.Enabled = true;
                this.docTemplateFilePath = dialog.FileName;
            }
        }

        /// <summary>
        /// 保存文档模板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveDocTemplateBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.docTemplateFilePath != "")
                {
                    this.dxRichTextBox.SaveDocument(this.docTemplateFilePath);
                }
                else
                {
                    SaveFileDialog dialog = new SaveFileDialog();
                    dialog.InitialDirectory = Common.DocTemplatePath;
                    dialog.Filter = "文档模板文件(*.doc)|*.doc";
                    dialog.FileName = "*.doc";
                    dialog.RestoreDirectory = true;
                    dialog.FilterIndex = 1;
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {

                        this.dxRichTextBox.SaveDocument(dialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           

        }

        /// <summary>
        /// 导出Doc文档
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exportDocBtn_Click(object sender, EventArgs e)
        {
            if (this.decimalsTextBox.Text.Length == 0)
            {
                MessageBox.Show("请输入有效位数");
                return;
            }
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.InitialDirectory = Common.AppPath;
            dialog.Filter = "文档文件(*.doc)|*.doc";
            dialog.FileName = "*.doc";
            if (dialog.ShowDialog() == DialogResult.OK)
            {

                try
                {
                    RichTextBox richTextBox = new RichTextBox();
                    richTextBox.Rtf = this.dxRichTextBox.RtfText;
                    this.exportDoc(dialog.FileName, richTextBox, this.decimalsTextBox.Text == "" ? -1 : Convert.ToInt32(this.decimalsTextBox.Text), this.isUnitComboBox.Text == "是" ? true : false);
                    MessageBox.Show("导出完成");
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }

        }
        #endregion

        
        
        private void docTemplateTreeView_AfterCheck(object sender, TreeViewEventArgs e)
        {
            PartProperties pro = (PartProperties)e.Node.Tag;
            if (e.Node.Nodes.Count != 0)
            {
                foreach (TreeNode n in e.Node.Nodes)
                {
                    n.Checked = e.Node.Checked;
                }
            }
            else
            {
                loadDocContextMenu(this.docTextContextMenu.Items);
            }

        }

        private void docTemplateRichTextBox_TextChanged(object sender, EventArgs e)
        {
            this.saveDocTemplateBtn.Enabled = true;
            this.exportDocBtn.Enabled = true;

        }

        #endregion

        #region 数据表格生成

        /// <summary>
        /// ExcelGrid列显示,CheckBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void partIDTableCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox box = (CheckBox)sender;
            switch (box.Tag.ToString())
            {
                case "allSelect":
                    foreach (Control c in this.panel2.Controls)
                    {
                        CheckBox temBox = c as CheckBox;

                        if (temBox != null)
                        {
                            if (temBox.Name == "isSumTableCheckBox")
                            {
                                continue;
                            }
                            temBox.Checked = box.Checked;
                        }
                    }
                    break;
                case "partInertiaMatrix_excel":
                    this.excelTemplateGridView.Columns["partInertiaMatrixIXX_excel"].Visible = box.Checked;
                    this.excelTemplateGridView.Columns["partInertiaMatrixIXY_excel"].Visible = box.Checked;
                    this.excelTemplateGridView.Columns["partInertiaMatrixIXZ_excel"].Visible = box.Checked;
                    this.excelTemplateGridView.Columns["partInertiaMatrixIYY_excel"].Visible = box.Checked;
                    this.excelTemplateGridView.Columns["partInertiaMatrixIYX_excel"].Visible = box.Checked;
                    this.excelTemplateGridView.Columns["partInertiaMatrixIYZ_excel"].Visible = box.Checked;
                    this.excelTemplateGridView.Columns["partInertiaMatrixIZX_excel"].Visible = box.Checked;
                    this.excelTemplateGridView.Columns["partInertiaMatrixIZY_excel"].Visible = box.Checked;
                    this.excelTemplateGridView.Columns["partInertiaMatrixIZZ_excel"].Visible = box.Checked;
                    break;
                case "partCenterOfGravity_excel":
                    this.excelTemplateGridView.Columns["partCenterOfGravityX_excel"].Visible = box.Checked;
                    this.excelTemplateGridView.Columns["partCenterOfGravityY_excel"].Visible = box.Checked;
                    this.excelTemplateGridView.Columns["partCenterOfGravityZ_excel"].Visible = box.Checked;
                    break;
                case "isSum":
                    if (box.Checked)
                    {
                        this.excelTemplateGridView.Rows.Add();
                        this.sumExcelTemplateGrid(this.excelTemplateGridView.Rows);
                    }
                    else
                    {
                        this.excelTemplateGridView.Rows.RemoveAt(this.excelTemplateGridView.Rows.Count - 1);
                    }
                    break;
                default:
                     this.excelTemplateGridView.Columns[box.Tag.ToString()].Visible = box.Checked;
                     break;
            }
        }

        #region Excel模板
        private string excelTemplateFilePath = ""; //Excel模板文件路径

        /// <summary>
        /// 新建Excel模板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void newExcelTemplatBtn_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("是否保存模板", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                saveExcelTemplateBtn_Click(null, null);
            }
            this.partIDTableCheckBox.Checked = true;
            this.parentIDTableCheckBox.Checked = true;
            this.partDensityTableCheckBox.Checked = true;
            this.partDimensionTableCheckBox.Checked = true;
            this.partQualityTableCheckBox.Checked = true;
            this.partAreaTableCheckBox.Checked = true;
            this.partCenterOfGravityTableCheckBox.Checked = true;
            this.partInertiaMatrixTableCheckBox.Checked = true;
            this.percentTableCheckBox.Checked = true;
            this.excelTemplateFilePath = "";
            this.excelTemplateNameLabel.Text = "";
            foreach(string customColumnText in this.customColumnComboBox.Items)
            {
                foreach (DataGridViewColumn column in this.excelTemplateGridView.Columns)
                {
                    if (column.Tag != null && column.Tag.ToString() == "customColumn" && column.HeaderText == customColumnText)
                    {
                        this.excelTemplateGridView.Columns.Remove(column);
                        break;
                    }
                }
            }
            this.customColumnComboBox.Items.Clear();
            this.customColumnComboBox.Text = "";
        }

        /// <summary>
        /// 打开Excel模板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openExcelTemplatBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = Common.ExcelTemplatePath;
            dialog.Filter = "模板文件(*.tpl)|*.tpl";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.excelTemplateFilePath = dialog.FileName;
                this.excelTemplateNameLabel.Text = Path.GetFileNameWithoutExtension(dialog.FileName);
                this.openExcelTemplateToXml(dialog.FileName);
                this.saveExcelTemplateBtn.Enabled = true;
            }
        }

        /// <summary>
        /// 保存Excel模板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveExcelTemplateBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.excelTemplateFilePath != "")
                {
                    this.saveExcelTemplateToXml(this.excelTemplateFilePath);
                    MessageBox.Show("保存成功");
                }
                else
                {
                    SaveFileDialog dialog = new SaveFileDialog();
                    dialog.InitialDirectory = Common.ExcelTemplatePath;
                    dialog.Filter = "模板文件(*.tpl)|*.tpl";
                    dialog.FileName = "*.tpl";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        this.saveExcelTemplateToXml(dialog.FileName);
                        this.excelTemplateFilePath = dialog.FileName;
                        this.excelTemplateNameLabel.Text = Path.GetFileName(dialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            

        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exportExcelBtn_Click(object sender, EventArgs e)
        {
            if (this.excelTemplatedecimalsTextBox.Text.Length == 0)
            {
                MessageBox.Show("请输入有效位数");
                return;
            }
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.InitialDirectory = Common.AppPath;
            dialog.Filter = "Excel文件(*.xls)|*.xls";
            dialog.FileName = "*.xls";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                DataTable dt = new DataTable();
                foreach (DataGridViewColumn column in this.excelTemplateGridView.Columns)
                {
                    if (column.Visible)
                    {
                        DataColumn dc = new DataColumn();
                        dc.Caption = column.HeaderText;
                        dc.ColumnName = column.Name;
                        dt.Columns.Add(dc);
                    }
                }


                foreach (DataGridViewRow row in this.excelTemplateGridView.Rows)
                {
                    int colIndex = 0;
                    DataRow dr = dt.NewRow();
                    dt.Rows.Add(dr);
                    foreach (DataGridViewColumn column in this.excelTemplateGridView.Columns)
                    {
                        if (column.Visible)
                        {
                            switch (column.HeaderText)
                            {
                                case "零部件密度(Kg/m^3)":
                                case "零部件体积（m^3）":
                                case "零部件质量(Kg)":
                                case "零部件面积(m^2)":
                                case "零部件重心X(m)":
                                case "零部件重心Y(m)":
                                case "零部件重心Z(m)":
                                case "零部件惯性矩阵IXX(Kg*m^2)":
                                case "零部件惯性矩阵IXY(Kg*m^2)":
                                case "零部件惯性矩阵IXZ(Kg*m^2)":
                                case "零部件惯性矩阵IYX(Kg*m^2)":
                                case "零部件惯性矩阵IYY(Kg*m^2)":
                                case "零部件惯性矩阵IYZ(Kg*m^2)":
                                case "零部件惯性矩阵IZX(Kg*m^2)":
                                case "零部件惯性矩阵IZY(Kg*m^2)":
                                case "零部件惯性矩阵IZZ(Kg*m^2)":

                                    dr[colIndex++] = Convert.ToDouble(row.Cells[column.Index].Value).ToString("e" + this.excelTemplatedecimalsTextBox.Text);
                                    break;
                                case "质量百分比(%)":
                                    string txt = row.Cells[column.Index].Value.ToString();
                                    txt = txt.Substring(0, txt.Length - 1);
                                    dr[colIndex++] = (Convert.ToDouble(txt).ToString("e"+this.excelTemplatedecimalsTextBox.Text))+"%";
                                    break;
                                default:
                                    dr[colIndex++] = row.Cells[column.Index].Value;
                                    break;
                            }
                        }
                    }
                }
                try
                {
                    this.exportExcelTemplateToExcel(dialog.FileName, dt);
                    MessageBox.Show("导出完成");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }
        #endregion

        #region 自定义列
        private int curSelectIndex = -1;//当前选择的自定义列项的索引值
        /// <summary>
        /// 增加自定义列
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void addColumnBtn_Click(object sender, EventArgs e)
        {
            if (this.customColumnComboBox.Text == "")
            {
                return;
            }
            this.customColumnComboBox.Items.Add(this.customColumnComboBox.Text);

            DataGridViewColumn column = new DataGridViewTextBoxColumn();
            column.Tag = "customColumn";
            column.HeaderText = this.customColumnComboBox.Text;
            this.excelTemplateGridView.Columns.Add(column);

            this.curSelectIndex = this.customColumnComboBox.Items.Count - 1;
        }

        private void modefyColmunBtn_Click(object sender, EventArgs e)
        {
            if (this.customColumnComboBox.Items.Count != 0)
            {
                foreach (DataGridViewColumn column in this.excelTemplateGridView.Columns)
                {
                    if (column.Tag != null && column.Tag.ToString() == "customColumn" && column.Visible && column.HeaderText == this.customColumnComboBox.Items[curSelectIndex].ToString())
                    {
                        column.HeaderText = this.customColumnComboBox.Text;
                        break;
                    }
                }
                this.customColumnComboBox.Items[curSelectIndex] = this.customColumnComboBox.Text;

            }
        }

        private void delColmunBtn_Click(object sender, EventArgs e)
        {
            if (this.customColumnComboBox.Items.Count == 0)
            {
                return;
            }
            this.customColumnComboBox.Items.Remove(this.customColumnComboBox.Text);

            foreach (DataGridViewColumn column in this.excelTemplateGridView.Columns)
            {
                if (column.Tag != null && column.Tag.ToString() == "customColumn" && column.HeaderText == this.customColumnComboBox.Text)
                {
                    this.excelTemplateGridView.Columns.Remove(column);
                    break;
                }
            }
            this.customColumnComboBox.SelectedIndex = this.customColumnComboBox.Items.Count - 1;
            this.customColumnComboBox.Text = "";

        }

        private void customColumnComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.curSelectIndex = this.customColumnComboBox.SelectedIndex;
        }
        #endregion

        private List<PartProperties> proList = new List<PartProperties>();
        private void excelTemplateTreeView_AfterCheck(object sender, TreeViewEventArgs e)
        {

            if (e.Node.Nodes.Count != 0)
            {
                foreach (TreeNode node in e.Node.Nodes)
                {
                    node.Checked = e.Node.Checked;
                }
            }
        }



        private void excelTemplateTreeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            List<PartProperties> proList = new List<PartProperties>();
            getCheckNodeData(this.excelTemplateTreeView.Nodes,proList);
            loadExcelDataTable(this.excelTemplateGridView.Rows, proList.ToArray());
            if (this.isSumTableCheckBox.Checked)
            {
                partIDTableCheckBox_CheckedChanged(this.isSumTableCheckBox, null);
            }
        }

        #endregion

        #region 图片生成

        private void exportPic_Click(object sender, EventArgs e)
        {
            if (this.partProperties == null)
            {
                return;
            }
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.InitialDirectory = Common.AppPath;
            dialog.Filter = "PNG文件(*.png)|*.png|BMP文件(*.bmp)|*.bmp|JPG文件(*.jpg)|*.jpg|GIF文件(*.gif)|*.gif";
            dialog.FilterIndex = 0;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ImageFormat format = null;
                switch (Path.GetExtension(dialog.FileName))
                {
                    case ".png":
                        format = ImageFormat.Png;
                        break;
                    case ".jpg":
                        format = ImageFormat.Jpeg;
                        break;
                    case ".gif":
                        format = ImageFormat.Gif;
                        break;
                    case ".bmp":
                        format = ImageFormat.Bmp;
                        break;
                }
                if (format != null)
                {
                    this.picChart.SaveImage(dialog.FileName, format);
                }

            }
        }

        private void chartTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox box = (ComboBox)sender;
            if (box.SelectedIndex == 0)
            {
                this.picChart.Series[0].ChartType = SeriesChartType.Column;
            }
            else if (box.SelectedIndex == 1)
            {
                this.picChart.Series[0].ChartType = SeriesChartType.Pie;
                //if (this.generatedPicturesTreeView.Nodes.Count != 0 && this.chartXComboBox.SelectedItem != null && this.chartYComboBox.SelectedItem != null)
                //{
                //    this.loadPieChart(this.generatedPicturesTreeView.Nodes[0].Nodes);
                //}
            }
            chartXComboBox_SelectedIndexChanged(null, null);
        }

        private void chartXComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //ComboBox box = (ComboBox)sender;
            if (this.generatedPicturesTreeView.Nodes.Count != 0 && this.chartXComboBox.SelectedItem != null && this.chartYComboBox.SelectedItem != null)
            {
                
                this.loadPicChart(this.generatedPicturesTreeView.Nodes[0].Nodes);
            }

        }

        private void generatedPicturesTreeView_AfterCheck(object sender, TreeViewEventArgs e)
        {
            TreeNode node = e.Node;
            if (node.Nodes.Count != 0)
            {
                foreach (TreeNode cn in node.Nodes)
                {
                    cn.Checked = node.Checked;
                }
            }
            else
            {
                if (this.chartXComboBox.SelectedItem != null && this.chartYComboBox.SelectedItem != null)
                {
                    this.loadPicChart(this.generatedPicturesTreeView.Nodes[0].Nodes);
                }
            }

        }
        #endregion


        private void picLabelCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.picChart.Series["Series1"].IsValueShownAsLabel=this.picLabelCheckBox.Checked;
        }


        private void decimalsTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            ToolStripTextBox box=(ToolStripTextBox)sender;
            if (!Char.IsNumber(e.KeyChar)&&e.KeyChar!=8)
            {
                e.Handled = true;//禁止执行,返回原字符
            }
            else
            {
                if (e.KeyChar!=8&&(Convert.ToInt32(box.Text + e.KeyChar) < 0 || Convert.ToInt32(box.Text + e.KeyChar) > 10))
                {
                    e.Handled = true;
                    return;
                }
                e.Handled = false;//允许执行,返回新字符.(此行可注释也行)
            }
        }

        private void picDecimalsTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox box = (TextBox)sender;
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != 8)
            {
                e.Handled = true;//禁止执行,返回原字符
            }
            else
            {
                if (e.KeyChar != 8 && (Convert.ToInt32(box.Text + e.KeyChar) < 0 || Convert.ToInt32(box.Text + e.KeyChar) > 10))
                {
                    e.Handled = true;
                    return;
                }
                e.Handled = false;//允许执行,返回新字符.(此行可注释也行)
            }
        }

        private void picDecimalsTextBox_TextChanged(object sender, EventArgs e)
        {
            chartXComboBox_SelectedIndexChanged(this.chartYComboBox, null);
        }

        private UserControlDXApplication.DXRichTextBox dxRichTextBox;
        private void tabBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabBox1.SelectedIndex == 1 && this.dxRichTextBox==null)
            {
                

                this.dxRichTextBox = new UserControlDXApplication.DXRichTextBox();
                this.groupBox4.Controls.Add(this.dxRichTextBox);
                this.dxRichTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
                this.dxRichTextBox.Location = new System.Drawing.Point(3, 17);
                this.dxRichTextBox.Name = "dxRichTextBox";
                //this.dxRichTextBox.RtfText = "";
                //this.dxRichTextBox.SelectText = "";
                this.dxRichTextBox.Size = new System.Drawing.Size(646, 450);
                this.dxRichTextBox.TabIndex = 1;

                if (this.partProperties != null)
                {
                    loadDocContextMenu(this.docTextContextMenu.Items);
                }
                
            }
        }







        
    }
}
