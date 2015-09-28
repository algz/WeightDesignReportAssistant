using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UserControlDXApplication
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            DXRichTextBox box = new DXRichTextBox();
            box.Dock = DockStyle.Fill;
            this.Controls.Add(box);
            PopupMenuObject pmo=new PopupMenuObject();
            pmo.Caption="自定义控件";
            pmo.onClickEvent = onCustomClick;
            //box.popupMenuObject = pmo;
        }

        private void onCustomClick(object sender, EventArgs e)
        {
            MessageBox.Show("Hello");
        }

        private void barMdiChildrenListItem1_ListItemClick(object sender, DevExpress.XtraBars.ListItemClickEventArgs e)
        {

        }

    }
}
