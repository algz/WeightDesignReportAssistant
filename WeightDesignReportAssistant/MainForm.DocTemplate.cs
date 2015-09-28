using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.ComponentModel;
using UserControlDXApplication;

namespace WeightDesignReportAssistant
{
    /// <summary>
    /// 质量特性获取
    /// </summary>
    partial class MainForm
    {
        /// <summary>
        /// 加载文档模板的质量特性TreeView
        /// </summary>
        /// <param name="parentID"></param>
        /// <param name="nodes"></param>
        private void loadDocTemplateTreeNode(string parentID, TreeNodeCollection nodes)
        {
            foreach (PartProperties pro in this.partProperties)
            {
                TreeNode node = new TreeNode();
                node.Name = pro.id;
                node.Text = pro.name;
                node.Tag = pro;
                node.Checked = true;
                nodes.Add(node);

                TreeNode[] cn = new TreeNode[20];
                for (int i = 0; i < cn.Length; i++)
                {
                    cn[i] = new TreeNode();
                    cn[i].Checked = true;
                }

                cn[0].Text = "ID:" + pro.id;
                cn[1].Text = "名称:" + pro.name;
                cn[2].Text = "父节点ID:" + pro.parentID;
                cn[3].Text = "密度:" + pro.density.ToString("E");
                cn[4].Text = "体积:" + pro.dimension.ToString("E");
                cn[5].Text = "质量:" + pro.quality.ToString("E");
                cn[6].Text = "面积:" + pro.area.ToString("E");
                cn[7].Text = "重心X:" + pro.centerOfGravityX.ToString("E");
                cn[8].Text = "重心Y:" + pro.centerOfGravityY.ToString("E");
                cn[9].Text = "重心Z:" + pro.centerOfGravityZ.ToString("E");
                cn[10].Text = "惯性矩阵IXX:" + pro.inertiaMatrixIXX.ToString("E");
                cn[11].Text = "惯性矩阵IXY:" + pro.inertiaMatrixIXY.ToString("E");
                cn[12].Text = "惯性矩阵IXZ:" + pro.inertiaMatrixIXZ.ToString("E");
                cn[13].Text = "惯性矩阵IYX:" + pro.inertiaMatrixIYX.ToString("E");
                cn[14].Text = "惯性矩阵IYY:" + pro.inertiaMatrixIYY.ToString("E");
                cn[15].Text = "惯性矩阵IYZ:" + pro.inertiaMatrixIYZ.ToString("E");
                cn[16].Text = "惯性矩阵IZX:" + pro.inertiaMatrixIZX.ToString("E");
                cn[17].Text = "惯性矩阵IZY:" + pro.inertiaMatrixIZY.ToString("E");
                cn[18].Text = "惯性矩阵IZZ:" + pro.inertiaMatrixIZZ.ToString("E");
                cn[19].Text = "质量百分比:" + pro.qualityPercent;

                node.Nodes.AddRange(cn);

            }
            
        }

        private void loadDocContextMenu(ToolStripItemCollection menus)
        {
            
            PopupMenuObject pmoRoot = new PopupMenuObject();
            pmoRoot.ChildMenu = new List<PopupMenuObject>();
            pmoRoot.Caption = "自定义数据";

            this.dxRichTextBox.popupMenuObject.Clear();
            this.dxRichTextBox.popupMenuObject.Add(pmoRoot);

            foreach (TreeNode n in this.docTemplateTreeView.Nodes)
            {
                PopupMenuObject pmo = new PopupMenuObject(n.Text);
                pmoRoot.ChildMenu.Add(pmo);

                //子结点
                foreach (TreeNode cn in n.Nodes)
                {
                    if (cn.Checked)
                    {
                        string txt=cn.Text.Split(':')[0];
                        PopupMenuObject pmoChil = new PopupMenuObject(txt);
                        pmoChil.Tag = " 『" + cn.Parent.Text + "∕" + txt + "』 ";//cn.Tag;
                        pmoChil.onClickEvent = new System.EventHandler(this.testToolStripMenuItem_Click);
                        pmo.ChildMenu.Add(pmoChil);
                    }

                }
            }

        }

        private void testToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PopupMenuObject pmobj = (PopupMenuObject)sender;


            this.dxRichTextBox.SelectText = pmobj.Tag.ToString();
            //this.docTemplateRichTextBox.SelectedText = ;
        }

        /// <summary>
        /// 导出Doc文档
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="richTextBox"></param>
        private void exportDoc(string fileName, RichTextBox richTextBox, int decimals,bool isUnit)
        {
            int statIndex = 0, endIndex = 0;
            bool v1 = false;
            while (richTextBox.Find("『") != -1)
            {
                statIndex = richTextBox.Find("『", statIndex, richTextBox.Text.Length, RichTextBoxFinds.None);
                endIndex = richTextBox.Find("』", statIndex, richTextBox.Text.Length, RichTextBoxFinds.None);
                richTextBox.SelectionStart = statIndex;
                richTextBox.SelectionLength = endIndex - statIndex + 1;

                string text = richTextBox.SelectedText.Trim();
                string[] s = text.Substring(1, text.Length - 2).Split(new char[] { '∕' });
                text = "";
                bool valid = true;
                foreach (PartProperties pro in this.partProperties)
                {
                    if (pro.name == s[0])
                    {
                        foreach (PropertyDescriptor des in TypeDescriptor.GetProperties(pro))
                        {
                            AttributeCollection attributes = des.Attributes;
                            DescriptionAttribute myAttribute = (DescriptionAttribute)attributes[typeof(DescriptionAttribute)];
                            if (myAttribute.Description.Contains(s[1]) || s[1].Contains(myAttribute.Description))
                            {
                                string txt= Convert.ToString(pro.GetType().GetProperty(des.Name).GetValue(pro, null));
                                //保留小数位数
                                if (decimals > -1)
                                {
                                    switch (myAttribute.Description)
                                    {
                                        case "零部件密度":
                                        case "零部件体积":
                                        case "零部件质量":
                                        case "零部件面积":
                                        case "零部件重心X":
                                        case "零部件重心Y":
                                        case "零部件重心Z":
                                        case "零部件惯性矩阵IXX":
                                        case "零部件惯性矩阵IXY":
                                        case "零部件惯性矩阵IXZ":
                                        case "零部件惯性矩阵IYX":
                                        case "零部件惯性矩阵IYY":
                                        case "零部件惯性矩阵IYZ":
                                        case "零部件惯性矩阵IZX":
                                        case "零部件惯性矩阵IZY":
                                        case "零部件惯性矩阵IZZ":
                                            txt=Convert.ToDouble(txt).ToString("e" + decimals);
                                            //txt=Math.Round(Convert.ToDouble(txt), decimals).ToString();
                                            break;
                                        case "质量百分比":
                                            txt=txt.Substring(0, txt.Length - 1);
                                            txt=Convert.ToDouble(txt).ToString("e" + decimals) + "%";
                                            break;
                                    }
                                }

                                //是否增加单位
                                if (isUnit)
                                {
                                    switch (myAttribute.Description)
                                    {
                                        case "零部件密度":
                                            txt += "kg/m³";
                                            break;
                                        case "零部件体积":
                                            txt += "m³";
                                            break;
                                        case "零部件质量":
                                            txt += "kg";
                                            break;
                                        case "零部件面积":
                                            txt += "m²";
                                            break;
                                        case "零部件重心X":
                                        case "零部件重心Y":
                                        case "零部件重心Z":
                                            txt += "m";
                                            break;
                                        case "零部件惯性矩阵IXX":
                                        case "零部件惯性矩阵IXY":
                                        case "零部件惯性矩阵IXZ":
                                        case "零部件惯性矩阵IYX":
                                        case "零部件惯性矩阵IYY":
                                        case "零部件惯性矩阵IYZ":
                                        case "零部件惯性矩阵IZX":
                                        case "零部件惯性矩阵IZY":
                                        case "零部件惯性矩阵IZZ":
                                            txt += "kg*m²";
                                            break;
                                    }
                                }
                                text = txt;
                                valid = false;
                                break;
                            }
                        }
                        //// Gets the attributes for the property.
                        //AttributeCollection attributes =TypeDescriptor.GetProperties(this.partProperties)["MyImage"].Attributes;
                        ///* Prints the description by retrieving the DescriptionAttribute 
                        // * from the AttributeCollection. */
                        //DescriptionAttribute myAttribute =(DescriptionAttribute)attributes[typeof(DescriptionAttribute)];
                        //Console.WriteLine(myAttribute.Description);

                        richTextBox.SelectedText = text;
                        break;
                    }

                }

                if (valid)
                {
                    richTextBox.SelectedText = "";
                    v1 = true;
                }
            }
            if (v1)
            {
                MessageBox.Show("数据不匹配");
            }
            richTextBox.SaveFile(fileName);

        }
    }
}
