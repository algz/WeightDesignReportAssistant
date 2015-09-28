using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.Utils.Menu;
using DevExpress.XtraRichEdit.Menu;

namespace UserControlDXApplication
{
    public partial class DXRichTextBox : UserControl
    {

        public DXRichTextBox()
        {
            InitializeComponent();
        }

        private void richEditControl1_PopupMenuShowing(object sender, DevExpress.XtraRichEdit.PopupMenuShowingEventArgs e)
        {
           // DXSubMenuItem subm = new DXSubMenuItem();
           // subm.Items.Add(new DXMenuItem("1"));
            
           // RichEditMenuItem m1 = (RichEditMenuItem)subm;
           // m1.
           //m1.Collection.Add(new RichEditMenuItem());
            foreach (PopupMenuObject pmo in this.popupMenuObject)
            {
                foreach (DXMenuItem m in e.Menu.Items)
                {
                    if (m.Caption == pmo.Caption)
                    {
                        e.Menu.Items.Remove(m);
                        break;
                    }
                }
            }

            if (this.popupMenuObject.Count!=0)
            {
                foreach (PopupMenuObject pmo in this.popupMenuObject)
                {
                    this.loadCustomPopupMenuItems(pmo, e.Menu.Items);
                }
                
            }
        }

        public void DXMenuItem_Click(object sender, EventArgs e)
        {
            DXMenuItem menu = (DXMenuItem)sender;
            
            PopupMenuObject pmobj = (PopupMenuObject)menu.Tag;
            if (pmobj.onClickEvent != null)
            {
                pmobj.onClickEvent(menu.Tag, e);
            }
            
        }

        /// <summary>
        /// 加载自定义弹出菜单项
        /// </summary>
        /// <param name="pmoList"></param>
        /// <param name="menuItems"></param>
        private void loadCustomPopupMenuItems(PopupMenuObject pmobj, DXMenuItemCollection menuItems)
        {
            DXMenuItem menu;

            if (pmobj.ChildMenu.Count == 0)
            {
                //一级子菜单(menu.Collection为null,并且不能set)
                menu = new DXMenuItem();
                menu.Click += new System.EventHandler(DXMenuItem_Click);// (pmobj.onClickEvent);

            }
            else
            {
                //多级子菜单(menu.Collection不为null)
                menu = new DXSubMenuItem();
                foreach (PopupMenuObject pmo in pmobj.ChildMenu)
                {
                    this.loadCustomPopupMenuItems(pmo, ((DXSubMenuItem)menu).Items);
                }
            }
            menu.Caption = pmobj.Caption;
            menu.Tag = pmobj;
            menuItems.Add(menu);
        }

        private void DXRichTextBox_Load(object sender, EventArgs e)
        {
            this.richEditControl1.RtfText = "";
        }

    }
}
