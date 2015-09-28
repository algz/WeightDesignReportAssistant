using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WeightDesignReportAssistant
{
    /// <summary>
    /// 质量特性获取
    /// </summary>
    partial class MainForm
    {
        /// <summary>
        /// 加载质量特性TreeView
        /// </summary>
        /// <param name="parentID"></param>
        /// <param name="nodes"></param>
        private void loadTreeNode(string parentID, TreeNodeCollection nodes)
        {
            foreach (PartProperties pro in this.partProperties)
            {
                if (pro.parentID == parentID)
                {
                    TreeNode node = new TreeNode();
                    node.Name = pro.id;
                    node.Text = pro.name;
                    nodes.Add(node);
                    loadTreeNode(pro.id, node.Nodes);
                }
            }
        }

        /// <summary>
        /// 加载质量特性GridView
        /// </summary>
        /// <param name="rows"></param>
        private void loadDataTable(DataGridViewRowCollection rows)
        {
            rows.Clear();
            foreach (PartProperties pro in this.partProperties)
            {
                int i = rows.Add();
                rows[i].Cells["partID"].Value = pro.id;
                rows[i].Cells["partName"].Value = pro.name;
                rows[i].Cells["parentID"].Value = pro.parentID;
                rows[i].Cells["partDensity"].Value = pro.density.ToString("e");
                rows[i].Cells["partDimension"].Value = pro.dimension.ToString("e");
                rows[i].Cells["partArea"].Value = pro.area.ToString("e");
                rows[i].Cells["partQuality"].Value = pro.quality.ToString("e");
                rows[i].Cells["partCenterOfGravityX"].Value = pro.centerOfGravityX.ToString("e");// String.Join(",", pro.centerOfGravity);
                rows[i].Cells["partCenterOfGravityY"].Value = pro.centerOfGravityY.ToString("e");
                rows[i].Cells["partCenterOfGravityZ"].Value = pro.centerOfGravityZ.ToString("e");
                rows[i].Cells["partInertiaMatrixIXX"].Value = pro.inertiaMatrixIXX.ToString("e");
                rows[i].Cells["partInertiaMatrixIXY"].Value = pro.inertiaMatrixIXY.ToString("e");
                rows[i].Cells["partInertiaMatrixIXZ"].Value = pro.inertiaMatrixIXZ.ToString("e");
                rows[i].Cells["partInertiaMatrixIYX"].Value = pro.inertiaMatrixIYX.ToString("e");
                rows[i].Cells["partInertiaMatrixIYY"].Value = pro.inertiaMatrixIYY.ToString("e");
                rows[i].Cells["partInertiaMatrixIYZ"].Value = pro.inertiaMatrixIYZ.ToString("e");
                rows[i].Cells["partInertiaMatrixIZX"].Value = pro.inertiaMatrixIZX.ToString("e");
                rows[i].Cells["partInertiaMatrixIZY"].Value = pro.inertiaMatrixIZY.ToString("e");
                rows[i].Cells["partInertiaMatrixIZZ"].Value = pro.inertiaMatrixIZZ.ToString("e");
            }
        }

        /// <summary>
        /// 获取的Catia数据加载到整个应用程序控件
        /// </summary>
        private void loadObtainCatia()
        {
            #region 质量特性获取
            this.partPropertiesTreeView.Nodes.Clear();
            loadTreeNode("0", this.partPropertiesTreeView.Nodes);
            this.partPropertiesTreeView.ExpandAll();
            loadDataTable(this.partPropertiesGridView.Rows);
            #endregion

            #region 文档段落生成
            this.docTemplateTreeView.Nodes.Clear();
            loadDocTemplateTreeNode("0", this.docTemplateTreeView.Nodes);
            this.docTemplateTreeView.ExpandAll();
            this.docTextContextMenu.Items.Clear();
            this.exportDocBtn.Enabled = true;

            if (this.dxRichTextBox != null)
            {
                loadDocContextMenu(this.docTextContextMenu.Items);
            }
            #endregion

            #region 数据表格生成
            this.excelTemplateTreeView.Nodes.Clear();
            loadExcelTemplateTreeNode("0", this.excelTemplateTreeView.Nodes);
            this.excelTemplateTreeView.ExpandAll();
            loadExcelDataTable(this.excelTemplateGridView.Rows, this.partProperties);
            this.exportExcelBtn.Enabled = true;
            #endregion

            #region 图片生成
            loadGeneratedPicturesTreeNode("0", this.generatedPicturesTreeView.Nodes);
            this.generatedPicturesTreeView.ExpandAll();
            this.exportPicBtn.Enabled = true;
            #endregion
        }
    
    }
}
