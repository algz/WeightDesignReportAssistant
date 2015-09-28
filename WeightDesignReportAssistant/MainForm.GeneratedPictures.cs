using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.ComponentModel;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing;

namespace WeightDesignReportAssistant
{
    /// <summary>
    /// 质量特性获取
    /// </summary>
    partial class MainForm
    {
        /// <summary>
        /// 加载生成图片的质量特性TreeView
        /// </summary>
        /// <param name="parentID"></param>
        /// <param name="nodes"></param>
        private void loadGeneratedPicturesTreeNode(string parentID, TreeNodeCollection nodes)
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
                    if (pro.id == "0" || pro.id == "1")
                    {
                        loadGeneratedPicturesTreeNode(pro.id, node.Nodes);
                    }
                    
                }
            }
        }

        //饼图   dt数据结构为 columndata(数据)  columnname(文本) 这两列  
        private void loadPieChart(TreeNodeCollection nodes)//(DataTable _dt, string _title)
        {

            if (nodes == null || nodes.Count == 0)
            {
                return;
            }
            this.picChart.Series["Series1"].Points.Clear();
            foreach (TreeNode node in nodes)
            {
                if (node.Checked)
                {
                    PartProperties pro = (PartProperties)node.Tag;
                    foreach (PropertyDescriptor des in TypeDescriptor.GetProperties(pro))
                    {
                        AttributeCollection attributes = des.Attributes;
                        DescriptionAttribute myAttribute = (DescriptionAttribute)attributes[typeof(DescriptionAttribute)];
                        if (this.chartXComboBox.SelectedItem != null && myAttribute.Description.Contains(this.chartXComboBox.SelectedItem.ToString()))
                        {
                            object o = pro.GetType().GetProperty(des.Name).GetValue(pro, null);
                            //dr["X"] = Convert.ToString(o);

                            int ptIdx = this.picChart.Series["Series1"].Points.AddY(Convert.ToString(o));
                            DataPoint pt = this.picChart.Series["Series1"].Points[ptIdx];
                            pt.LegendText = Convert.ToString(o);// dr["columnname"].ToString() + " " + "#PERCENT{P2}" + " [ " + "#VAL{D} 人" + " ]";//右边标签列显示的文字  
                            pt.Label = Convert.ToString(o); //dr["columnname"].ToString() + " " + "#PERCENT{P2}" + " [ " + "#VAL{D} 人" + " ]"; //圆饼外显示的信息  
                        }

                    }
                }
            }
            //this.picChart.Series["Series1"].LegendText= this.chartYComboBox.Text;

        }  

        /// <summary>
        /// 
        /// </summary>
        /// <param name="nodes"></param>
        private void loadPicChart(TreeNodeCollection nodes)
        {
            if (nodes == null || nodes.Count == 0)
            {
                return;
            }

            this.picChart.Series["Series1"].Points.Clear();

            DataTable dt = new DataTable();
            DataColumn dc = new DataColumn();
            dc.ColumnName = "X";
            dt.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Y";
            dt.Columns.Add(dc);

            foreach (TreeNode node in nodes)
            {
                if (node.Checked)
                {
                    DataRow dr = dt.NewRow();
                    PartProperties pro = (PartProperties)node.Tag;
                    foreach (PropertyDescriptor des in TypeDescriptor.GetProperties(pro))
                    {
                        AttributeCollection attributes = des.Attributes;
                        DescriptionAttribute myAttribute = (DescriptionAttribute)attributes[typeof(DescriptionAttribute)];
                        if (this.chartXComboBox.SelectedItem != null && myAttribute.Description.Contains(this.chartXComboBox.SelectedItem.ToString()))
                        {
                            object o = pro.GetType().GetProperty(des.Name).GetValue(pro, null);
                            dr["X"] = Convert.ToString(o);

                        }
                        if (this.chartYComboBox.SelectedItem != null && myAttribute.Description.Contains(this.chartYComboBox.SelectedItem.ToString()))
                        {
                            object o = pro.GetType().GetProperty(des.Name).GetValue(pro, null);
                            if (this.picDecimalsTextBox.Text != "")
                            {
                                //switch (myAttribute.Description)
                                //{
                                //    case "零部件密度":
                                //    case "零部件体积":
                                //    case "零部件质量":
                                //    case "零部件面积":
                                //    case "零部件重心X":
                                //    case "零部件重心Y":
                                //    case "零部件重心Z":
                                //    case "零部件惯性矩阵IXX":
                                //    case "零部件惯性矩阵IXY":
                                //    case "零部件惯性矩阵IXZ":
                                //    case "零部件惯性矩阵IYX":
                                //    case "零部件惯性矩阵IYY":
                                //    case "零部件惯性矩阵IYZ":
                                //    case "零部件惯性矩阵IZX":
                                //    case "零部件惯性矩阵IZY":
                                //    case "零部件惯性矩阵IZZ":
                                //        o = Convert.ToDouble(o).ToString("e" + this.picDecimalsTextBox.Text);
                                //        break;
                                //    case "质量百分比":
                                //        o = o.ToString().Substring(0, o.ToString().Length - 1);
                                //        o = Convert.ToDouble(o).ToString("e" + this.picDecimalsTextBox.Text)+ "%";
                                //        break;
                                //}
                                
                            }
                            dr["Y"] = Convert.ToString(o);
                        }
                    }
                    dt.Rows.Add(dr);
                }
            }

            DataView dv = dt.DefaultView;
            this.picChart.Series["Series1"].Points.DataBindXY(dv, "X", dv, "Y");
            if (this.picChart.Series["Series1"].ChartType == SeriesChartType.Column)
            {
                this.picChart.Series["Series1"].LegendText = this.chartYComboBox.Text;
            }
            else
            {
                this.picChart.Series["Series1"].LegendText = "";
            }
            

        }
    }
}
