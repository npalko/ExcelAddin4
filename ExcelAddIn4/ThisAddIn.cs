/* pg 702
 * 
 * 
 * DESIGNER QUESTIONS
 *  - does the designer automatically add hooks for winform events into its
 *  region?
 * hooks for windows forms events: in parital class (*.designer.cs)?
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
 */


using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelAddIn4
{
    public partial class ThisAddIn
    {

        // pg 203 - named ranges
        // pg 219 - ranges

        // 
        // 

        // Application.StatusBar = string

        UserControl1 control;
        public Microsoft.Office.Tools.CustomTaskPane pane;
        //Button button;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            control = new UserControl1();

            //button = new Button();
            //button.Text = "Hello";
            // button.Text = Application.ActiveWorkbook.name
            //control.Controls.Add(button);

            pane = CustomTaskPanes.Add(control, "my pane");
            pane.Visible = true;



            //Excel.Range r1 = Application.get_Range("A1", missing);
            //r1.Value2 = 8;


            // Application.WindowActivate += 
            //  new Excel.AppEvents_WindowActivateEventHandler(
            //  Application_WindowActivate)
        }

        public void setCell(string cell, int value)
        {
            Excel.Range r = Application.get_Range(cell, missing);
            r.Value2 = value;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        private void button_Click(object sender, System.EventArgs e)
        {
        }

 //       private void Application_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
 //       {
            // changes button text whenever a different workbook is selected
            // within the excel instance

 //           button.Text = Wb.Name;
//        }

        // to add to internal startup:
        // this.button.Click += new System.EventHanlder(this.button_Click);

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
