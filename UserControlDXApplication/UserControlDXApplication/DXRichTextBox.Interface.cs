using DevExpress.XtraRichEdit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace UserControlDXApplication
{
    public partial class DXRichTextBox : UserControl
    {
        /// <summary>
        /// 编辑区弹出菜单子菜单项对象
        /// </summary>
        //[CategoryAttribute("自定义属性"),
        //DescriptionAttribute("标题1"),
        //DefaultValue("曲线")]
        public List<PopupMenuObject> popupMenuObject=new List<PopupMenuObject>();

        #region 属性

        public override string Text
        {
            get
            {
                return this.richEditControl1.Text;
            }
            set
            {
                this.richEditControl1.Text = value;
            }
        }

        public string RtfText
        {
            get
            {
                return this.richEditControl1.RtfText;
            }
            set
            {
                this.richEditControl1.RtfText = value;
            }
        }

        public string SelectText
        {
            get{
                return this.richEditControl1.Document.Selection.ToString();
            }
            set{
                this.richEditControl1.Document.InsertText(this.richEditControl1.Document.CaretPosition, value);
        
            }
        }

        #endregion

        #region 方法

        /// <summary>
        /// 打开文档
        /// </summary>
        /// <param name="file"></param>
        public void OpenDocument(string file)
        {
            //string filter = "Word2003(*.doc)|*.doc|Word2007(*.docx)|*.docx|RTF(*.rtf)|*.rtf|HTM(*.htm)|*.htm|HTML(*.html)|*.html|All File(*.*)|*.*";
            //string file = FileDialogHelper.Open("打开文件", filter);
            if (!string.IsNullOrEmpty(file))
            {
                //string htmlContent = File.ReadAllText(file, Encoding.Default);
                //this.richEditControl1.HtmlText = htmlContent;
                string path = Path.GetFullPath(file);
                string extension = Path.GetExtension(file);
                switch (extension.ToLower())
                {
                    case ".htm":
                    case ".html":
                        this.richEditControl1.Document.LoadDocument(file, DocumentFormat.Html, path);
                        break;
                    case ".doc":
                        this.richEditControl1.Document.LoadDocument(file, DocumentFormat.Doc, path);
                        break;
                    case ".docx":
                        this.richEditControl1.Document.LoadDocument(file, DocumentFormat.OpenXml, path);
                        break;
                    case ".rtf":
                        this.richEditControl1.Document.LoadDocument(file, DocumentFormat.Rtf, path);
                        break;
                    default:
                        this.richEditControl1.Document.LoadDocument(file, DocumentFormat.PlainText, path);
                        break;
                }

                //DocumentRange range = richEditControl1.Document.Range;
                //CharacterProperties cp = this.richEditControl1.Document.BeginUpdateCharacters(range);
                //cp.FontName = "新宋体";
                //cp.FontSize = 12;
                //this.richEditControl1.Document.EndUpdateCharacters(cp);
            }
        }

        /// <summary>
        /// 关闭文档
        /// </summary>
        /// <param name="fileName"></param>
        public void SaveDocument(string fileName)
        {
            this.richEditControl1.SaveDocument(fileName, DocumentFormat.Doc);
        }

        #endregion

    }
}
