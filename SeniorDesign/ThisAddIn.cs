using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Diagnostics;

namespace SeniorDesign
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            
        
            }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        static public string context_Doc()
        {
            string textFromDoc = Globals.ThisAddIn.Application.ActiveDocument.Range().Text;
            string text = "";
            text += textFromDoc;
            Debug.WriteLine( "testing",text);
            return text;
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
