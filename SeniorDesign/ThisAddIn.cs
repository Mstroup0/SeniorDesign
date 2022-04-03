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
using System.Runtime.InteropServices;

namespace SeniorDesign
{
    public partial class ThisAddIn
    {
        bool IsDatasetDirty { get; set; }
        TrainedDataSet dataSet { get; set; }
        public IEnumerable<string> words;

        private KeyboardHook hook1 = new KeyboardHook();
        private KeyboardHook hook2 = new KeyboardHook();
        private KeyboardHook hook3 = new KeyboardHook();
        private KeyboardHook hook4 = new KeyboardHook();
 

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            dataSet = new TrainedDataSet();
            hook1.KeyPressed += new EventHandler<KeyPressedEventArgs>(addWord1);
            hook1.RegisterHotKey(1, Keys.D1); // register the alt + num combination as hot key.
            //Alt = 1,
            //Control = 2,
            //Shift = 4,
            //Win = 8

            hook2.KeyPressed += new EventHandler<KeyPressedEventArgs>(addWord2);
            hook2.RegisterHotKey(1, Keys.D2);

            hook3.KeyPressed += new EventHandler<KeyPressedEventArgs>(addWord3);
            hook3.RegisterHotKey(1, Keys.D3);

            hook4.KeyPressed += new EventHandler<KeyPressedEventArgs>(addWord4);
            hook4.RegisterHotKey(1, Keys.D4);
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void addWord1(object sender, KeyPressedEventArgs e)
        {
            //Debug.WriteLine("Pressed");
            string suggestion = arrayWords(0);
            PUPrintWord(suggestion);
        }

        void addWord2(object sender, KeyPressedEventArgs e)
        {
            string suggestion = arrayWords(1);
            PUPrintWord(suggestion);
        }

        void addWord3(object sender, KeyPressedEventArgs e)
        {
            string suggestion = arrayWords(2);
            PUPrintWord(suggestion);
        }

        void addWord4(object sender, KeyPressedEventArgs e)
        {
            string suggestion = arrayWords(3);
            PUPrintWord(suggestion);
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
            if (cursorPos != 0)
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
                end = cursorPos + 36;
            }
            //Set the start and end of the selection range
            startO = start;
            endO = end;

            //gets the selections and inputs as a string 
            textFromDoc = Globals.ThisAddIn.Application.ActiveDocument.Range(ref startO, ref endO).Text.ToString();
            text += textFromDoc;

            //test printing selection
            Debug.WriteLine("Selections Testing: ", text);


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

            char[] delimiter = { ' ', ',', '.', '?', '!', '\n' };
            string[] words = wordsRange2.Split(delimiter);
            foreach (var word in words)
            {
                Debug.WriteLine("testing Doc var.last:" + word);
            }
            words = words.Where(x => !string.IsNullOrEmpty(x)).ToArray();
            foreach (var word in words)
            {
                Debug.WriteLine("testing Doc var.last:" + word);
            }
            // gets the last word in the range
            string lastWord = words.Last().ToString();
            Debug.WriteLine("testing Doc var.last:" + lastWord);

            // get rid of any white space
            string noWhite = String.Concat(lastWord.Where(c => !Char.IsWhiteSpace(c)));
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
            char[] delimiter = { ' ', ',', '.', '?', '!', '\n' };
            string[] words = wordsRange2.Split(delimiter);
            foreach (var word in words)
            {
                Debug.WriteLine("testing Doc var.last:" + word);
            }
            words = words.Where(x => !string.IsNullOrEmpty(x)).ToArray();
            foreach (var word in words)
            {
                Debug.WriteLine("testing Doc var.last:" + word);
            }
            int size = words.Length;
            // gets the last word in the range
            string last2ndWord = words[size - 3].ToString();
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
            string noWlast2Word;

            string lastWord = GetLastWord();
            Debug.WriteLine("testing Doc:" + lastWord);

            string noWlastWord = String.Concat(lastWord.Where(c => !Char.IsWhiteSpace(c)));

            IEnumerable<string> suggestedWords1 = dataSet.Next4Words(noWlastWord, 8);


            if (!suggestedWords1.Any())
            {
                string last2Word = GetLastWord();
                Debug.WriteLine("testing Doc:" + last2Word);

                noWlast2Word = String.Concat(last2Word.Where(c => !Char.IsWhiteSpace(c)));

                suggestedWords1 = dataSet.Next4Words(noWlast2Word, 8);

            }
            else
            {
                noWlast2Word = "";
            }
            checkDictionary(noWlastWord, noWlast2Word);

            List<string> wordsAsList = suggestedWords1.ToList();
            wordsAsList.Remove("{{end}}");

            foreach (string word in suggestedWords1)
            {
                Debug.WriteLine("4 Suggested word:" + word);
            }

            words = wordsAsList.AsEnumerable();
        }

        public void checkDictionary(String word1In, String word2In)
         {
            
           string word1 = word1In.TryToLower();
            string word2 = word2In.TryToLower();
            List<string> W12 = new List<string>();
            
                if (word2 != null)
                {
                    W12.Add(word2);
                     W12.Add(word1);
                    dataSet.TrainS(W12);
                }
                else
                {
                    W12.Add("");
                    W12.Add(word1);
                    dataSet.TrainS(W12);
                }
         }


        //Train(List<string> sentence)
        private void OpenDataSet()
        {
            if (AskIfSaveFirst())
            {
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

    public sealed class KeyboardHook : IDisposable
    {
        // Registers a hot key with Windows.
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        // Unregisters the hot key with Windows.
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // Represents the window that is used internally to get the messages.
        private class Window : NativeWindow, IDisposable
        {
            private static int WM_HOTKEY = 0x0312;

            public Window()
            {
                // create the handle for the window.
                this.CreateHandle(new CreateParams());
            }

            // Overridden to get the notifications.
            protected override void WndProc(ref Message m)
            {
                base.WndProc(ref m);

                // check if we got a hot key pressed.
                if (m.Msg == WM_HOTKEY)
                {
                    // get the keys.
                    Keys key = (Keys)(((int)m.LParam >> 16) & 0xFFFF);
                    ModifierKeys modifier = (ModifierKeys)((int)m.LParam & 0xFFFF);

                    // invoke the event to notify the parent.
                    if (KeyPressed != null)
                        KeyPressed(this, new KeyPressedEventArgs(modifier, key));
                }
            }

            public event EventHandler<KeyPressedEventArgs> KeyPressed;

            #region IDisposable Members

            public void Dispose()
            {
                this.DestroyHandle();
            }

            #endregion
        }

        private Window _window = new Window();
        private int _currentId;

        public KeyboardHook()
        {
            // register the event of the inner native window.
            _window.KeyPressed += delegate (object sender, KeyPressedEventArgs args)
            {
                if (KeyPressed != null)
                    KeyPressed(this, args);
            };
        }

        // Registers a hot key in the system.
        public void RegisterHotKey(uint modifier, Keys key)
        {
            // increment the counter.
            _currentId++;

            // register the hot key.
            if (!RegisterHotKey(_window.Handle, _currentId, (uint)modifier, (uint)key))
                throw new InvalidOperationException("Couldn’t register the hot key.");
        }

        // A hot key has been pressed.
        public event EventHandler<KeyPressedEventArgs> KeyPressed;

        #region IDisposable Members

        public void Dispose()
        {
            // unregister all the registered hot keys.
            for (int i = _currentId; i > 0; i--)
            {
                UnregisterHotKey(_window.Handle, i);
            }

            // dispose the inner native window.
            _window.Dispose();
        }

        #endregion
    }

    // Event Args for the event that is fired after the hot key has been pressed.
    public class KeyPressedEventArgs : EventArgs
    {
        private ModifierKeys _modifier;
        private Keys _key;

        internal KeyPressedEventArgs(ModifierKeys modifier, Keys key)
        {
            _modifier = modifier;
            _key = key;
        }
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
