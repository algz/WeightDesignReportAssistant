using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Collections;
using Excel=Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;

namespace WeightDesignReportAssistant
{
    /// <summary>
    /// 质量特性获取
    /// </summary>
    partial class MainForm
    {
        /// <summary>
        /// 保存Excel模板
        /// </summary>
        /// <param name="fileName"></param>
        private void saveExcelTemplateToXml(string fileName)
        {
            XElement root = new XElement("tpl");

            XElement colElement = new XElement("表头列");
            root.Add(colElement);

            foreach (DataGridViewColumn column in this.excelTemplateGridView.Columns)
            {
                if (column.Visible)
                {
                    colElement.Add(new XElement("列名称", column.HeaderText));
                }
            }

            root.Add(new XElement("是否总计", this.isSumTableCheckBox.Checked));

            root.Save(fileName);
        }


        private void clearCustomColumn()
        {
            //清除全部自定义表格
            //while (this.customColumnComboBox.Items.Count != 0)
            //{
            //    this.customColumnComboBox.SelectedIndex = this.customColumnComboBox.Items.Count - 1;
            //    this.customColumnComboBox.Text = this.customColumnComboBox.SelectedItem.ToString();
            //    this.delColmunBtn_Click(null, null);
            //}
            while (true)
            {
                DataGridViewColumn column = this.excelTemplateGridView.Columns[this.excelTemplateGridView.Columns.Count-1];
                if (column.Tag != null && column.Tag.ToString() == "customColumn")
                {
                    this.excelTemplateGridView.Columns.Remove(column);
                }
                if (column.Tag == null)
                {
                    break;
                }

            }
            this.customColumnComboBox.Items.Clear();
        }

        /// <summary>
        /// 打开Excel模板
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private void openExcelTemplateToXml(string fileName)
        {
            clearCustomColumn();

            XDocument doc = XDocument.Load(fileName);

            this.isSumTableCheckBox.Checked = Convert.ToBoolean(doc.Descendants("是否总计").First().Value);

            this.partIDTableCheckBox.Checked = false;
            this.parentIDTableCheckBox.Checked = false;
            this.partDensityTableCheckBox.Checked = false;
            this.partDimensionTableCheckBox.Checked = false;
            this.partQualityTableCheckBox.Checked = false;
            this.partAreaTableCheckBox.Checked = false;
            this.partCenterOfGravityTableCheckBox.Checked = false;
            this.partInertiaMatrixTableCheckBox.Checked = false;
            this.percentTableCheckBox.Checked = false;

            var elements = from c in doc.Descendants("列名称") select c;
            for (int i = 0; i < elements.Count(); i++)
            {
                XElement element = ((XElement)elements.ElementAt(i));
                if (element.Value == "零部件ID")
                {
                    this.partIDTableCheckBox.Checked =true;
                }
                else if (element.Value == "零部件父节点ID")
                {
                    this.parentIDTableCheckBox.Checked = true;
                }
                else if (element.Value == "零部件密度(Kg/m^3)")
                {
                    this.partDensityTableCheckBox.Checked = true;
                }
                else if (element.Value == "零部件体积（m^3）")
                {
                    this.partDimensionTableCheckBox.Checked = true;
                }
                else if (element.Value == "零部件质量(Kg)")
                {
                    this.partQualityTableCheckBox.Checked = true;
                }
                else if (element.Value == "零部件面积(m^2)")
                {
                    this.partAreaTableCheckBox.Checked = true;
                }
                else if (element.Value == "零部件重心X(m)" || element.Value == "零部件重心Y(m)" || element.Value == "零部件重心Z(m)")
                {
                    this.partCenterOfGravityTableCheckBox.Checked = true;
                }
                else if (element.Value == "零部件转动惯量IXX(Kg*m^2)" || element.Value == "零部件转动惯量IXY(Kg*m^2)" || element.Value == "零部件转动惯量IXZ(Kg*m^2)" ||
                    element.Value == "零部件转动惯量IYX(Kg*m^2)" || element.Value == "零部件转动惯量IYY(Kg*m^2)" || element.Value == "零部件转动惯量IYZ(Kg*m^2)" ||
                    element.Value == "零部件转动惯量IZX(Kg*m^2)" || element.Value == "零部件转动惯量IZY(Kg*m^2)" || element.Value == "零部件转动惯量IZZ(Kg*m^2)")
                {
                    this.partInertiaMatrixTableCheckBox.Checked = true;
                }
                else if (element.Value == "质量百分比(%)")
                {
                    this.percentTableCheckBox.Checked = true;
                }
                else if (element.Value == "零部件名称")
                {
                    continue;
                }else
                {
                    this.customColumnComboBox.Text = element.Value;
                    addColumnBtn_Click(null, null);
                }
            }
        }

        /// <summary>
        /// 加载文档模板的质量特性TreeView
        /// </summary>
        /// <param name="parentID"></param>
        /// <param name="nodes"></param>
        private void loadExcelTemplateTreeNode(string parentID, TreeNodeCollection nodes)
        {
            foreach (PartProperties pro in this.partProperties)
            {
                if (pro.parentID == parentID)
                {
                    TreeNode node = new TreeNode();
                    node.Name = pro.id;
                    node.Text = pro.name;
                    node.Tag = pro;
                    node.Checked = true;
                    nodes.Add(node);
                    loadExcelTemplateTreeNode(pro.id, node.Nodes);
                }
            }
        }

        /// <summary>
        /// 加载Excel数据表格GridView
        /// </summary>
        /// <param name="rows"></param>
        private void loadExcelDataTable(DataGridViewRowCollection rows,PartProperties[] pros)
        {
            rows.Clear();
            foreach (PartProperties pro in pros)
            {
                int i = rows.Add();
                rows[i].Cells["partID_excel"].Value = pro.id;
                rows[i].Cells["partName_excel"].Value = pro.name;
                rows[i].Cells["parentID_excel"].Value = pro.parentID;
                rows[i].Cells["partDensity_excel"].Value = pro.density.ToString("e");
                rows[i].Cells["partDimension_excel"].Value = pro.dimension.ToString("e");
                rows[i].Cells["partArea_excel"].Value = pro.area.ToString("e");
                rows[i].Cells["partQuality_excel"].Value = pro.quality.ToString("e");
                rows[i].Cells["partPercent_excel"].Value = pro.qualityPercent;
                rows[i].Cells["partCenterOfGravityX_excel"].Value = pro.centerOfGravityX.ToString("e");//String.Join(",", pro.centerOfGravity);
                rows[i].Cells["partCenterOfGravityY_excel"].Value = pro.centerOfGravityY.ToString("e");
                rows[i].Cells["partCenterOfGravityZ_excel"].Value = pro.centerOfGravityZ.ToString("e");
                rows[i].Cells["partInertiaMatrixIXX_excel"].Value = pro.inertiaMatrixIXX.ToString("e");
                rows[i].Cells["partInertiaMatrixIXY_excel"].Value = pro.inertiaMatrixIXY.ToString("e");
                rows[i].Cells["partInertiaMatrixIXZ_excel"].Value = pro.inertiaMatrixIXZ.ToString("e");
                rows[i].Cells["partInertiaMatrixIYY_excel"].Value = pro.inertiaMatrixIYY.ToString("e");
                rows[i].Cells["partInertiaMatrixIYX_excel"].Value = pro.inertiaMatrixIYX.ToString("e");
                rows[i].Cells["partInertiaMatrixIYZ_excel"].Value = pro.inertiaMatrixIYZ.ToString("e");
                rows[i].Cells["partInertiaMatrixIZX_excel"].Value = pro.inertiaMatrixIZX.ToString("e");
                rows[i].Cells["partInertiaMatrixIZY_excel"].Value = pro.inertiaMatrixIZY.ToString("e");
                rows[i].Cells["partInertiaMatrixIZZ_excel"].Value = pro.inertiaMatrixIZZ.ToString("e");
            }
        }

        private void sumExcelTemplateGrid(DataGridViewRowCollection rows)
        {
            float[] val = { 0, 0, 0, 0 };
            foreach (DataGridViewRow row in rows)
            {
                val[0] += Convert.ToSingle(row.Cells["partDensity_excel"].Value);
                val[1] += Convert.ToSingle(row.Cells["partDimension_excel"].Value);
                val[2] += Convert.ToSingle(row.Cells["partArea_excel"].Value);
                val[3] += Convert.ToSingle(row.Cells["partQuality_excel"].Value);
            }
            rows[rows.Count - 1].Cells["partDensity_excel"].Value = val[0];
            rows[rows.Count - 1].Cells["partDimension_excel"].Value = val[1];
            rows[rows.Count - 1].Cells["partArea_excel"].Value = val[2];
            rows[rows.Count - 1].Cells["partQuality_excel"].Value = val[3];
        }

        private void getCheckNodeData(TreeNodeCollection nodes, List<PartProperties> prosList)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Checked)
                {
                    prosList.Add((PartProperties)node.Tag);
                }
                if (node.Nodes.Count != 0)
                {
                    getCheckNodeData(node.Nodes, prosList);
                }


            }
        }

        private void exportExcelTemplateToExcel(string fileName,DataTable dt)
        {
            //创建Excel对象  
            Excel.Application excelApp = new Excel.Application();
            //新建工作簿  
            Excel.Workbook workBook = excelApp.Workbooks.Add(true);
            //新建工作表  
            Excel.Worksheet worksheet = workBook.ActiveSheet as Excel.Worksheet;

            try
            {
                //设置Excel可见  
                excelApp.Visible = false;

                ////设置标题  
                //Excel.Range titleRange = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, columnCount]);//选取单元格  
                //titleRange.Merge(true);//合并单元格  
                //titleRange.Value2 = strTitle; //设置单元格内文本  
                //titleRange.Font.Name = "宋体";//设置字体  
                //titleRange.Font.Size = 18;//字体大小  
                //titleRange.Font.Bold = false;//加粗显示  
                //titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//水平居中  
                //titleRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;//垂直居中  
                //titleRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//设置边框  
                //titleRange.Borders.Weight = Excel.XlBorderWeight.xlThin;//边框常规粗细  

                //设置表头  
                //int nMax = 9;
                //int nMin = 4;

                //int rowCount = nMax - nMin + 1;//总行数  
                //const int columnCount = 4;//总列数  
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    DataColumn column = dt.Columns[i];
                    //Excel.Range headRange = worksheet.Cells[2, i + 1] as Excel.Range;//获取表头单元格  

                    Excel.Range headRange = worksheet.Cells[1, i + 1] as Excel.Range;//获取表头单元格,不用标题则从1开始  
                    headRange.Value2 = column.Caption;//设置单元格文本  
                    headRange.Font.Name = "宋体";//设置字体  
                    headRange.Font.Size = 12;//字体大小  
                    headRange.Font.Bold = false;//加粗显示  
                    headRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//水平居中  
                    headRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;//垂直居中  
                    headRange.ColumnWidth = 20;//设置列宽  
                    headRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//设置边框  
                    headRange.Borders.Weight = Excel.XlBorderWeight.xlThin;//边框常规粗细  

                }

                ////设置每列格式  
                //for (int i = 0; i < columnCount; i++)
                //{
                //    //Excel.Range contentRange = worksheet.get_Range(worksheet.Cells[3, i + 1], worksheet.Cells[rowCount - 1 + 3, i + 1]);  
                //    Excel.Range contentRange = worksheet.get_Range(worksheet.Cells[2, i + 1], worksheet.Cells[rowCount - 1 + 3, i + 1]);//不用标题则从第二行开始  
                //    contentRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//水平居中 
                //    contentRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;//垂直居中  
                //    //contentRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//设置边框 
                //    // contentRange.Borders.Weight = Excel.XlBorderWeight.xlThin;//边框常规粗细  
                //    contentRange.WrapText = true;//自动换行  
                //    contentRange.NumberFormatLocal = "@";//文本格式  
                //}

                int[] num = { 0, 0, 0 };
                double[] sum = { 0.1, 0.1, 0.1 };
                //填充数据  
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        DataColumn column = dt.Columns[j];
                        worksheet.Cells[i + 2, j + 1] = dt.Rows[i][j];

                        if (this.isSumTableCheckBox.Checked && (column.Caption == "零部件体积" || column.Caption == "零部件质量" || column.Caption == "零部件面积"))
                        {
                            double val = Convert.ToDouble(dt.Rows[i][j]);
                            switch (column.Caption)
                            {
                                case "零部件体积":
                                    num[0] = j + 1;
                                    sum[0] = sum[0] + val;
                                    break;
                                case "零部件质量":
                                    num[1] = j + 1;
                                    sum[1] = sum[1] + val;
                                    break;
                                case "零部件面积":
                                    num[2] = j + 1;
                                    sum[2] = sum[2] + val;
                                    break;
                            }
                        }
                    }
                }
                if (this.isSumTableCheckBox.Checked)
                {
                    for (int i = 0; i < 3; i++)
                    {
                        if (num[i] != 0)
                        {
                            worksheet.Cells[dt.Rows.Count + 2, num[i]] = sum[i];
                        }
                        
                    }
                        

                }
            }
            catch(Exception e)
            {
                throw e;
            }
            finally
            {
                workBook.SaveCopyAs(fileName);
                workBook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                excelApp.Quit();

                IntPtr t = new IntPtr(excelApp.Hwnd);
                int k = 0;
                Common.GetWindowThreadProcessId(t, out   k);
                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
                p.Kill();

                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workBook);
                Marshal.ReleaseComObject(excelApp);
                worksheet = null;
                workBook = null;
                excelApp = null;
                //GC.Collect();System.Runtime.InteropServices
            }

        }  
    }
}
