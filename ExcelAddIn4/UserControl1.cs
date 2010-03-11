using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelAddIn4
{
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            Globals.ThisAddIn.setCell("A1", trackBar1.Value);
        }

        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            Globals.ThisAddIn.setCell("B1", trackBar2.Value);
        }

        private void trackBar3_Scroll(object sender, EventArgs e)
        {
            Globals.ThisAddIn.setCell("C1", trackBar3.Value);
        }

        private void trackBar4_Scroll(object sender, EventArgs e)
        {
            Globals.ThisAddIn.setCell("D1", trackBar4.Value);
        }

        private void trackBar5_Scroll(object sender, EventArgs e)
        {
            Globals.ThisAddIn.setCell("E1", trackBar5.Value);
        }


    }
}
