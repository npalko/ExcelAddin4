/* pg 702
 * 
 * 

 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * Globals class is used to interact with Ribbon
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * Architecture:
 * ThisAddIn
 *      CustomTaskPlane
 *          UserControl1 (sliders)
 *      Ribbon
 * 
 * 
 * 
 * Ah-Hahs:
 * * double cick on winform components to generate events
 * 
 * 
 * Code Snips:
 * 
 * Application.StatusBar = string
 * 
 * Working with ranges (pg 203, 219):
 *  Excel.Range r1 = Application.get_Range("A1", missing);
 *  r1.Value2 = 8;
 * 
 */


using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelAddIn4
{
    public partial class ThisAddIn
    {

        Microsoft.Office.Tools.CustomTaskPane pane;
        ModelControl modelControl;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            modelControl = new ModelControl();
            pane = CustomTaskPanes.Add(modelControl, "ExcelAddIn4");


            Application.WindowActivate +=
                new Excel.AppEvents_WindowActivateEventHandler(
                Application_WindowActivate);
        }
        public void PaneVisibleToggle()
        {
            pane.Visible = !pane.Visible;
        }


        public void setCell(string cell, int value)
        {
            Excel.Range r = Application.get_Range(cell, missing);
            r.Value2 = value;
        }



        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Application_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            // called whenever a different workbook is selected
            modelControl.UpdateWorkbookName(Wb.Name);

            // whever we switch worksheets, move the worksheet values into
            // the toggle
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
