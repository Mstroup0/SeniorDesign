using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using WordPredictionLibrary.Core;
using System.Windows.Forms;
using System.IO;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Input;
using System.Windows.Documents;
using Microsoft.Office.Interop.Word;

namespace SeniorDesign
{
    public partial class ThisAddIn
    {
        bool IsDatasetDirty { get; set; }
        TrainedDataSet dataSet { get; set; }
        public IEnumerable<string> words;

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
            Debug.WriteLine("testing Doc:" + docCon);
            string docCon2 = docCon;
            Debug.WriteLine("testing Doc:" + docCon2);
            //string docConLast = docCon2.Split(new char[] { ' ', ',', '.', '?', '!', '\n'}).Last();
            string docConLast2 = String.Concat(docCon2.Where(c => !Char.IsWhiteSpace(c)));
            string lastW =  docConLast2;

            Debug.WriteLine("testing Doc:" + lastW);
            string suggestedWord = dataSet.SuggestNext(lastW);

            Debug.WriteLine("1 Suggested word:" + suggestedWord);
            IEnumerable<string> suggestedWords = dataSet.Next4Words(lastW, 4);
            words = suggestedWords;

            string suggests = " ";
            foreach (string word in suggestedWords)
            {
                Debug.WriteLine("4 Suggested word:" + word);
            }

           // docConT += " " + suggestedWord;
           // docConT += " " + suggests;

           // objPare.Range.Text += docConT;
          
        }

        private void OpenDataSet()
        {
            if (AskIfSaveFirst())
            {
                //string selectedFile = ShowFileDialog(openFileDialog);
                string selectedFile = "C:\\Users\\kuro0\\Source\\Repos\\Mstroup0\\SeniorDesign\\SeniorDesign\\Texts\\Dictionary.txt";
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

        public void PUPrintWord(string suggestion) => PrintWord(suggestion);
        private void PrintWord(string suggestion) // Prints word at current possition
        {
            Word.Selection currentSelection = Application.Selection;

            // Store the user's current Overtype selection
            bool userOvertype = Application.Options.Overtype;

            // Make sure Overtype is turned off.
            if (Application.Options.Overtype)
            {
                Application.Options.Overtype = false;
            }

            // Test to see if selection is an insertion point.
            if (currentSelection.Type == Word.WdSelectionType.wdSelectionIP)
            {
                currentSelection.TypeText(suggestion);
                currentSelection.TypeParagraph();
            }
            else
                if (currentSelection.Type == Word.WdSelectionType.wdSelectionNormal)
            {
                // Move to start of selection.
                if (Application.Options.ReplaceSelection)
                {
                    object direction = Word.WdCollapseDirection.wdCollapseStart;
                    currentSelection.Collapse(ref direction);
                }
                currentSelection.TypeText(suggestion);
                currentSelection.TypeParagraph();
            }
            else
            {
                // Do nothing.
            }

            // Restore the user's Overtype selection
            Application.Options.Overtype = userOvertype;
        }

        private void CursorPos1() // possiblility 2
        {

            Object wordObject = null;
            Microsoft.Office.Interop.Word.Application word = null;
            Document document = null;

            try
            {
                wordObject = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");

                word = (Microsoft.Office.Interop.Word.Application)wordObject;
                word.Visible = false;
                word.ScreenUpdating = false;
                string fullPath = word.ActiveDocument.FullName;

                document = word.ActiveDocument;

                int count = document.Words.Count;
                for (int k = 1; k <= count; k++)
                {
                    string text = document.Words[k].Text;
                    MessageBox.Show(text);
                }

                if (document.Paragraphs.Count > 0)
                {
                    var paragraph = document.Paragraphs.First;
                    var lastCharPos = paragraph.Range.Sentences.First.End - 1;
                    MessageBox.Show(lastCharPos.ToString());
                }
            }
            catch (Exception ex) { 
            MessageBox.Show(ex.ToString());
                }
            }
           
        private void SaveDataSet()
        {
            string selectedFile = "C:\\Users\\kuro0\\Source\\Repos\\Mstroup0\\SeniorDesign\\SeniorDesign\\Texts\\Dictionary.txt";
            if (!string.IsNullOrWhiteSpace(selectedFile))
            {
                if (TrainedDataSet.SerializeToXml(dataSet, selectedFile))
                {
                    IsDatasetDirty = false;
                }
            }
        }
        public string arrayWords(int pos)
        {
            IEnumerable<string> suggestedWords = words;
            return suggestedWords.ElementAt(pos);
        }
        public IEnumerable<string> UpdateLabels()
        {
            return GetSuggestion();
        }
        private IEnumerable<string> GetSuggestion()
        {
            return words;
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


/* OG print code
  private void SelectionInsertText1() 
    {
       Word.Selection currentSelection = Application.Selection;
       bool userOvertype = Application.Options.Overtype;
        if (Application.Options.Overtype)
        {
            Application.Options.Overtype = false;
        }

        if (currentSelection.Type == Word.WdSelectionType.wdSelectionIP)
        {
            currentSelection.TypeText("Inserting at insertion point. ");
            currentSelection.TypeParagraph();
        }
        else
            if (currentSelection.Type == Word.WdSelectionType.wdSelectionNormal)
        {
            
            if (Application.Options.ReplaceSelection)
            {
                object direction = Word.WdCollapseDirection.wdCollapseStart;
                currentSelection.Collapse(ref direction);
            }
            currentSelection.TypeText("Inserting before a text block. ");
            currentSelection.TypeParagraph();
        }
        else
        {
            Do nothing.
        }

 
        Application.Options.Overtype = userOvertype;
*/
