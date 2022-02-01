using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using WordPredictionLibrary.Core;
using System.Windows.Forms;
using System.IO;
using System;
using System.Linq;
using System.Collections.Generic;

namespace SeniorDesign
{
    public partial class ThisAddIn
    {
        bool IsDatasetDirty { get; set; }
        TrainedDataSet dataSet { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            dataSet = new TrainedDataSet();
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

            var docCon =  Context_Doc();
            string docConT = docCon;

            string docCon2 = docCon;
            docCon2 = String.Concat(docCon.Where(c => !Char.IsWhiteSpace(c)));
            var lastW =  docCon2;

            Debug.Write("testing Doc:" + lastW);
            string suggestedWord = dataSet.SuggestNext(lastW);

            Debug.WriteLine("Suggested word:" + suggestedWord);
            IEnumerable<string> suggestedWords = dataSet.Next4Words(lastW, 4);

            string suggests = " ";
            foreach (string word in suggestedWords)
            {
                suggests += " " + word;
                Debug.WriteLine("Suggested word:" + suggests);
            }

           // docConT += " " + suggestedWord;
            docConT += " " + suggests;

            objPare.Range.Text += docConT;
        }

        private void OpenDataSet()
        {
            if (AskIfSaveFirst())
            {
                //string selectedFile = ShowFileDialog(openFileDialog);
                string selectedFile = "C:\\Users\\kuro0\\source\\repos\\SeniorDesign\\SeniorDesign\\Texts\\Dictionary.txt";
                Debug.WriteLine("file " + selectedFile);
                if (!string.IsNullOrWhiteSpace(selectedFile) && File.Exists(selectedFile))
                {
                    dataSet = TrainedDataSet.DeserializeFromXml(selectedFile);
                    if (dataSet != null)
                    {
                        OnDataSetLoaded();
                    }
                }
            }
        }

        private void OnDataSetLoaded()
        {
            IsDatasetDirty = false;
        }

        private bool AskIfSaveFirst()
        {
            if (dataSet != null && dataSet.TotalSampleSize > 1)
            {
                if (IsDatasetDirty)
                {
                    DialogResult result = MessageBox.Show("Do you wish to save current Trained Data Set?", "Save?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    switch (result)
                    {
                        case DialogResult.Yes:
                            SaveDataSet();
                            break;

                        case DialogResult.No:
                            break;

                        case DialogResult.Cancel:
                        default:
                            return false;
                    }
                }
            }
            return true;
        }

        private void SaveDataSet()
        {
            string selectedFile = "C:\\Users\\kuro0\\source\\repos\\SeniorDesign\\SeniorDesign\\Texts\\Dictionary.txt";
            if (!string.IsNullOrWhiteSpace(selectedFile))
            {
                if (TrainedDataSet.SerializeToXml(dataSet, selectedFile))
                {
                    IsDatasetDirty = false;
                }
            }
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
