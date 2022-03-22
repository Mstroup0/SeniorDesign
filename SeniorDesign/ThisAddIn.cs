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

        private string GetLastWordsinRange()
        {
            //cursor starting position
            int cursorPos = Application.Selection.Start;
            Debug.WriteLine("Testing starting postion ", cursorPos);

            //Variables
            string text = "";
            string textFromDoc = "";
            int start, end;
            object startO, endO;

            //Finds the range based off of the text cursors position 
            if (cursorPos != 0 )
            {
                if ((cursorPos - 36) > 0)
                {
                    start = cursorPos - 36;
                }
                else
                {
                    start = 0;
                }
                end = cursorPos;
                
            }
            else
            {
                start = cursorPos;
                end = cursorPos;
            }
            //Set the start and end of the selection range
            startO = start;
            endO = end;

            //gets the selections and inputs as a string 
            textFromDoc = Globals.ThisAddIn.Application.ActiveDocument.Range(ref startO, ref endO).Text;
            text += textFromDoc;

            //test printing selection
            Debug.WriteLine( "Selections Testing: ",text);
            
            
            //Returns the last word
            return text;


        }
        public string GetLastWord()
        {
            string lWord;

            //calls for the string in the range
            string wordsRange = GetLastWordsinRange();
            Debug.WriteLine("testing words in range/getlastword:" + wordsRange);
            
            // set to another string to keep the og
            string wordsRange2 = wordsRange;
            Debug.WriteLine("testing Doc:" + wordsRange2);
            
            var words = wordsRange2.Split( ' ', ',', '.', '?', '!', '\n' );

            // gets the last word in the range
            string lastWord = words.Last().ToString();
            Debug.WriteLine("testing Doc var.last:" + lastWord);

            // get rid of any white space
            string noWhite= String.Concat(lastWord.Where(c => !Char.IsWhiteSpace(c)));
            Debug.WriteLine("testing Doc:" + noWhite);
            //Sets the last Word
            lWord = noWhite;

            return lWord;
        }
        public string Get2ndLastWord()
        {
            string lWord;

            //calls for the string in the range
            string wordsRange = GetLastWordsinRange();
            Debug.WriteLine("testing words in range/getlastword:" + wordsRange);

            // set to another string to keep the og
            string wordsRange2 = wordsRange;
            Debug.WriteLine("testing Doc:" + wordsRange2);

            var words = wordsRange2.Split(' ', ',', '.', '?', '!', '\n');

            int size = words.Length;
            // gets the last word in the range
            string last2ndWord = words[ size - 2].ToString();
            Debug.WriteLine("testing Doc var.last:" + last2ndWord);

            // get rid of any white space
            string noWhite = String.Concat(last2ndWord.Where(c => !Char.IsWhiteSpace(c)));
            Debug.WriteLine("testing Doc:" + noWhite);
            //Sets the last Word
            lWord = noWhite;

            return lWord;
        }



        public void Suggest()
        {
            OpenDataSet();

            string lastWord =  GetLastWord();
            Debug.WriteLine("testing Doc:" + lastWord);

            string noWlastWord = String.Concat(lastWord.Where(c => !Char.IsWhiteSpace(c)));

            string suggestedWord = dataSet.SuggestNext(noWlastWord);
            Debug.WriteLine("1 Suggested word:" + suggestedWord);
            
            IEnumerable<string> suggestedWords = dataSet.Next4Words(lastWord, 4);
            words = suggestedWords;

            if (!suggestedWords.Any())
            { 
                string last2Word = GetLastWord();
                Debug.WriteLine("testing Doc:" + last2Word);

                string noWlast2Word = String.Concat(last2Word.Where(c => !Char.IsWhiteSpace(c)));

                string suggested2Word = dataSet.SuggestNext(noWlast2Word);
                Debug.WriteLine("1 Suggested word:" + suggested2Word);
                suggestedWords = dataSet.Next4Words(last2Word, 4);

            }
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
                //string selectedFile = "C:\\Users\\kuro0\\Source\\Repos\\Mstroup0\\SeniorDesign\\SeniorDesign\\Texts\\Dictionary.txt";
                string selectedFile = Environment.GetEnvironmentVariable("PREDICTION_DICTIONARY", EnvironmentVariableTarget.Machine);
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

        // Prints the suggested word
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

        private void SaveDataSet()
        {
            //string selectedFile = "C:\\Users\\kuro0\\Source\\Repos\\Mstroup0\\SeniorDesign\\SeniorDesign\\Texts\\Dictionary.txt";
            string selectedFile = Environment.GetEnvironmentVariable("PREDICTION_DICTIONARY", EnvironmentVariableTarget.Machine);
            if (!string.IsNullOrWhiteSpace(selectedFile))
            {
                if (TrainedDataSet.SerializeToXml(dataSet, selectedFile))
                {
                    IsDatasetDirty = false;
                }
            }
        }

        //Gets Suggested word at specific position
        public string arrayWords(int pos)
        {
            IEnumerable<string> suggestedWords = words;
            return suggestedWords.ElementAt(pos);
        }
        //Gets the suggestions
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
