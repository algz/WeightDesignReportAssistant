using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace WeightDesignReportAssistant
{
    public class PartProperties
    {

        public PartProperties()
        {
            this.id="1";
            this.name = "name1";
            this.parentID = "0";
            this.density=0.1f;
            this.dimension=0.1f;
            this.quality = 0.1f;
            this.area = 0.1f;
            this.centerOfGravityX = 0.1F;
            this.centerOfGravityY = 0.2f;
            this.centerOfGravityZ = 0.3f;
            this.inertiaMatrixIXX = 0.1f;
            this.inertiaMatrixIXY = 0.2f;
            this.inertiaMatrixIXZ = 0.3f;
            this.inertiaMatrixIYY = 0.4f;
            this.inertiaMatrixIYX = 0.5f;
            this.inertiaMatrixIYZ = 0.6f;
            this.inertiaMatrixIZX = 0.7f;
            this.inertiaMatrixIZY = 0.8f;
            this.inertiaMatrixIZZ = 0.9f;
        }

        /// <summary>
        /// 1	零部件ID	字符串	用于记录描述零/部件的唯一编码
        /// </summary>
        [Description("零部件ID")] 
        public string id
        {
            get;
            set;
        }

        /// <summary>
        /// 2	零部件名称	字符串	用于记录零/部件的名称
        /// </summary>
        [Description("零部件名称")] 
        public string name
        {
            get;
            set;
        }

        /// <summary>
        /// 3	零部件父节点ID	字符串	用于记录该零/部件父级节点ID
        /// </summary>
        [Description("零部件父节点ID")] 
        public string parentID
        {
            get;
            set;
        }

        /// <summary>
        /// 4	零部件密度	浮点	用于记录该零/部件材料密度
        /// （若部件由多种材料组成，该项为-1）
        /// </summary>
        [Description("零部件密度")] 
        public double density
        {
            get;
            set;
        }

        /// <summary>
        /// 5	零部件体积	浮点	用于记录该零/部件体积
        ///（部件为所有下属零件体积求和）
        /// </summary>
        [Description("零部件体积")]
        public double dimension
        {
            get;
            set;
        }

        /// <summary>
        /// 6 零部件质量	浮点	用于记录该零/部件质量
        /// （部件为所有下属零件质量求和）
        /// </summary>
        [Description("零部件质量")]
        public double quality
        {
            get;
            set;
        }

        [Description("质量百分比")]
        public string qualityPercent
        {
            get;
            set;
        }

        /// <summary>
        /// 7	零部件面积	浮点	用于记录该零/部件面积
        ///（部件为所有下属零件面积求和）
        /// </summary>
        [Description("零部件面积")]
        public double area 
        {
            get;
            set;
        }

        /// <summary>
        /// 8	零部件重心	浮点数组（3）	用于记录零/部件的重心X,Y,Z
        ///（三个坐标分量）
        /// </summary>
        [Description("零部件重心X")]
        public double centerOfGravityX
        {
            get;
            set;
        }

        [Description("零部件重心Y")]
        public double centerOfGravityY
        {
            get;
            set;
        }

        [Description("零部件重心Z")]
        public double centerOfGravityZ
        {
            get;
            set;
        }

        /// <summary>
        /// 9	零部件惯性矩阵	浮点数组（6）	用于记录零/部件的惯性数据IXX,IYY,IZZ,IXY,IYZ,IZX
        ///（三个惯性矩，三个惯性积）
        /// </summary>
        [Description("零部件惯性矩阵IXX")]
        public double inertiaMatrixIXX
        {
            get;
            set;
        }

        [Description("零部件惯性矩阵IXY")]
        public double inertiaMatrixIXY
        {
            get;
            set;
        }

        [Description("零部件惯性矩阵IXZ")]
        public double inertiaMatrixIXZ
        {
            get;
            set;
        }

        [Description("零部件惯性矩阵IYY")]
        public double inertiaMatrixIYY
        {
            get;
            set;
        }

        [Description("零部件惯性矩阵IYX")]
        public double inertiaMatrixIYX
        {
            get;
            set;
        }

        [Description("零部件惯性矩阵IYZ")]
        public double inertiaMatrixIYZ
        {
            get;
            set;
        }

        [Description("零部件惯性矩阵IZX")]
        public double inertiaMatrixIZX
        {
            get;
            set;
        }

        [Description("零部件惯性矩阵IZY")]
        public double inertiaMatrixIZY
        {
            get;
            set;
        }

        [Description("零部件惯性矩阵IZZ")]
        public double inertiaMatrixIZZ
        {
            get;
            set;
        }
    }
}
