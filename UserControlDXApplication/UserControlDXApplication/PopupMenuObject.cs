using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UserControlDXApplication
{
    public class PopupMenuObject
    {
        public PopupMenuObject()
        {
            this.ChildMenu = new List<PopupMenuObject>();
        }

        public PopupMenuObject(string text):this()
        {
            this.Caption = text;
        }

        /// <summary>
        /// 文本标题
        /// </summary>
        public string Caption
        {
            get;
            set;
        }

        /// <summary>
        /// 单击事件
        /// </summary>
        public EventHandler onClickEvent
        {
            get;
            set;
        }

        /// <summary>
        /// 自定义对象
        /// </summary>
        public object Tag
        {
            get;
            set;
        }

        public List<PopupMenuObject> ChildMenu
        {
            get;
            set;
        }
    }
}
