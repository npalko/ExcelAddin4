﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn4
{
    public partial class Ribbon1 : OfficeRibbon
    {
        public Ribbon1()
        {
            InitializeComponent();
        }
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {       
        }
        private void togglePane_Click(object sender, RibbonControlEventArgs e)
        {
            // had to double-click on button in ribbon designer to
            // auto-generate event handler in Ribbon1.Designer.cs

            // is this the best way to handle the event?

            bool currentState = Globals.ThisAddIn.pane.Visible;
            Globals.ThisAddIn.pane.Visible = !currentState;

        }
    }
}
