using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Xml.Linq;
using System.Windows.Forms;
using System.Collections;
using System.Runtime.InteropServices;

namespace WeightDesignReportAssistant
{
    class Common
    {
        //应用程序路径
        public static string AppPath = System.AppDomain.CurrentDomain.BaseDirectory;

        //文档模板路径
        public static string DocTemplatePath = AppPath + @"Template\doc";

        //数据表格模板路径
        public static string ExcelTemplatePath = AppPath + @"Template\excel";

        /// <summary>
        /// 初始化静态变量
        /// </summary>
        static Common()
        {
            if (!Directory.Exists(DocTemplatePath))
            {
                Directory.CreateDirectory(DocTemplatePath);
            }
            if (!Directory.Exists(ExcelTemplatePath))
            {
                Directory.CreateDirectory(ExcelTemplatePath);
            }
        }

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out   int ID);   

        public static void savePWDToXml(string fileName, PartProperties[] properties)
        {
            XElement root = new XElement("pwd");
            
            foreach (PartProperties property in properties)
            {
                XElement element = new XElement("零部件",
                    new XElement("零部件ID", property.id),
                    new XElement("零部件名称", property.name),
                    new XElement("零部件父节点ID", property.parentID),
                    new XElement("零部件密度", property.density),
                    new XElement("零部件体积", property.dimension),
                    new XElement("零部件质量", property.quality),
                    new XElement("零部件面积", property.area),
                    new XElement("零部件重心X", property.centerOfGravityX),
                    new XElement("零部件重心Y",  property.centerOfGravityY),
                    new XElement("零部件重心Z", property.centerOfGravityZ),
                    new XElement("零部件转动惯量IXX", property.inertiaMatrixIXX),
                    new XElement("零部件转动惯量IXY", property.inertiaMatrixIXY),
                    new XElement("零部件转动惯量IXZ", property.inertiaMatrixIXZ),
                    new XElement("零部件转动惯量IYX", property.inertiaMatrixIYX),
                    new XElement("零部件转动惯量IYY", property.inertiaMatrixIYY),
                    new XElement("零部件转动惯量IYZ", property.inertiaMatrixIYZ),
                    new XElement("零部件转动惯量IZX", property.inertiaMatrixIZX),
                    new XElement("零部件转动惯量IZY", property.inertiaMatrixIZY),
                    new XElement("零部件转动惯量IZZ", property.inertiaMatrixIZZ)
                    );
                root.Add(element);
            }
            root.Save(fileName);
        }

        /// <summary>
        /// 转换项目XML文件到本地对象数组this.PartProperties[]
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static PartProperties[] ConvertPWDXmlToPartPropertiesArray(string fileName)
        {
            XDocument doc = XDocument.Load(fileName);
            var elements = from c in doc.Descendants("零部件") select c;

            PartProperties[] pros = new PartProperties[elements.Count()];
            try
            {

                for (int i = 0; i < elements.Count(); i++)
                {
                    pros[i] = new PartProperties();
                    XElement element = ((XElement)elements.ElementAt(i));
                    pros[i].id = element.Element("零部件ID").Value;
                    pros[i].name = element.Element("零部件名称").Value;
                    pros[i].parentID = element.Element("零部件父节点ID").Value;
                    pros[i].density = Convert.ToDouble(element.Element("零部件密度").Value);
                    pros[i].dimension = Convert.ToDouble(element.Element("零部件体积").Value);
                    pros[i].quality = Convert.ToDouble(element.Element("零部件质量").Value);
                    pros[i].qualityPercent = Math.Round(pros[i].quality / pros[0].quality * 100, 4) + "%";
                    pros[i].area = Convert.ToDouble(element.Element("零部件面积").Value);
                    pros[i].centerOfGravityX = Convert.ToDouble(element.Element("零部件重心X").Value);
                    pros[i].centerOfGravityY = Convert.ToDouble(element.Element("零部件重心Y").Value);
                    pros[i].centerOfGravityZ = Convert.ToDouble(element.Element("零部件重心Z").Value);
                    pros[i].inertiaMatrixIXX = Convert.ToDouble(element.Element("零部件转动惯量IXX").Value);
                    pros[i].inertiaMatrixIXY = Convert.ToDouble(element.Element("零部件转动惯量IXY").Value);
                    pros[i].inertiaMatrixIXZ = Convert.ToDouble(element.Element("零部件转动惯量IXZ").Value);
                    pros[i].inertiaMatrixIYX = Convert.ToDouble(element.Element("零部件转动惯量IYX").Value);
                    pros[i].inertiaMatrixIYY = Convert.ToDouble(element.Element("零部件转动惯量IYY").Value);
                    pros[i].inertiaMatrixIYZ = Convert.ToDouble(element.Element("零部件转动惯量IYZ").Value);
                    pros[i].inertiaMatrixIZX = Convert.ToDouble(element.Element("零部件转动惯量IZX").Value);
                    pros[i].inertiaMatrixIZY = Convert.ToDouble(element.Element("零部件转动惯量IZY").Value);
                    pros[i].inertiaMatrixIZZ = Convert.ToDouble(element.Element("零部件转动惯量IZZ").Value);
                }
            }
            catch
            {
                pros = null;
            }
            
            return pros;
        }

        public static PartProperties[] ConvertCatiaInterfaceToPartPropertiesArray(List<CatiaCommon.PartProperties> lstGravity)
        {
            if (lstGravity.Count() == 0)
            {
                return null;
            }

            try
            {
                PartProperties[] partProperties = new PartProperties[lstGravity.Count()];

                //转换PartProperties.vb==>PartProperties.cs
                partProperties = new PartProperties[lstGravity.Count];

                for (int i = 0; i < lstGravity.Count; i++)
                {
                    CatiaCommon.PartProperties cpp = lstGravity[i];
                    partProperties[i] = new PartProperties();
                    partProperties[i].id = cpp.id;
                    partProperties[i].name = cpp.name;
                    partProperties[i].parentID = cpp.parentID;
                    partProperties[i].density = cpp.density;
                    partProperties[i].dimension = cpp.dimension;
                    partProperties[i].quality = cpp.quality;
                    partProperties[i].area = cpp.area;
                    partProperties[i].centerOfGravityX = cpp.centerOfGravity[0];
                    partProperties[i].centerOfGravityY = cpp.centerOfGravity[1];
                    partProperties[i].centerOfGravityZ = cpp.centerOfGravity[2];
                    partProperties[i].inertiaMatrixIXX = cpp.inertiaMatrix[0];
                    partProperties[i].inertiaMatrixIXY = cpp.inertiaMatrix[1];
                    partProperties[i].inertiaMatrixIXZ = cpp.inertiaMatrix[2];
                    partProperties[i].inertiaMatrixIYX = cpp.inertiaMatrix[3];
                    partProperties[i].inertiaMatrixIYY = cpp.inertiaMatrix[4];
                    partProperties[i].inertiaMatrixIYZ = cpp.inertiaMatrix[5];
                    partProperties[i].inertiaMatrixIZX = cpp.inertiaMatrix[6];
                    partProperties[i].inertiaMatrixIZY = cpp.inertiaMatrix[7];
                    partProperties[i].inertiaMatrixIZZ = cpp.inertiaMatrix[8];
                    partProperties[i].qualityPercent = (partProperties[i].quality / partProperties[0].quality * 100) + "%";
                }
                return partProperties;
            }
            catch
            {
                return null;
            }
            
        }

        /// <summary>
        /// 获取指定控件下的所有单选框控件
        /// </summary>
        /// <param name="container"></param>
        /// <param name="controlArrayList"></param>
        public static void GetCheckBoxs(Control container, ref ArrayList controlArrayList)
        {
            if (container.Controls.Count != 0)
            {
                foreach (Control c in container.Controls)
                {
                    GetCheckBoxs(c, ref controlArrayList);
                }
            }
            else
            {
                if (container is CheckBox)
                {
                    controlArrayList.Add(container);
                }
            }
            return;
        }
    }
}
