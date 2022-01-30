using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Diagnostics;
using Microsoft.Office.Tools.Ribbon;
using WordPredictionLibrary.Core;
using System.IO;
using System.Diagnostics;

namespace SeniorDesign
{
    public partial class ThisAddIn
    {
        bool IsDatasetDirty { get; set; }
        TrainedDataSet dataSet { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
         }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private string Context_Doc()
        {
            string textFromDoc = Globals.ThisAddIn.Application.ActiveDocument.Range().Text;
            string text = "";
            text += textFromDoc;
           // Debug.WriteLine( "testing",text);
            return text;
        }
        public void Suggest()
        {
            OpenDataSet();
            Microsoft.Office.Interop.Word._Document oDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Paragraph objPare;
            objPare = oDoc.Paragraphs.Add();

            string docCon = "";
            docCon += " " + Context_Doc();


            //error Is occuring here the lastword will not display in the debug output.
            // data is disappearing
            // Once fix words should predict. 
            string last = "";
            last += String.Concat( Context_Doc());

            Debug.Write("testing Doc: ", docCon);
            Debug.WriteLine("testing2: ", last);
            //objPare.Range.Text = lastWord; Testing
            string suggestedWord = "";
            suggestedWord += dataSet.SuggestNext(docCon);

            Debug.WriteLine("Suggested word: ", suggestedWord);
            Debug.WriteLine("testing Doc: ", docCon);

            IEnumerable<string> suggestedWords = dataSet.Next4Words(last, 4);

            docCon += " " + suggestedWord;

            //return suggestedWord;
        }

        private void OpenDataSet()
        {
            //if (AskIfSaveFirst())
            //{
            string selectedFile = "C:\\Users\\kuro0\\source\\repos\\SeniorDesign\\SeniorDesign\\Texts\\Dictionary.txt";
            dataSet = TrainedDataSet.DeserializeFromXml(selectedFile);

            //}
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
