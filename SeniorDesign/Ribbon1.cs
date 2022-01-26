using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordPredictionLibrary.Core;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Diagnostics;

namespace SeniorDesign
{
    public partial class Ribbon1
    {

        bool IsDatasetDirty { get; set; }
        TrainedDataSet dataSet { get; set; }



        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {


        }
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            OpenDataSet();
            Microsoft.Office.Interop.Word._Document oDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Paragraph objPare;
            objPare = oDoc.Paragraphs.Add();

            string docCon = "";
            docCon += " " + ThisAddIn.context_Doc();
            Debug.Write("testing Doc: ", docCon);


            //error Is occuring here the lastword will not display in the debug output.
            // data is disappearing
            // Once fix words should predict. 
            var lastWord = "";
            Debug.Write("testing Doc: ", docCon);
            lastWord += "" + docCon;
            Debug.Write("testing Doc: ", docCon);

            Debug.Write("testing2: ", lastWord);

            Debug.Write("testing Doc: ", docCon);
            //objPare.Range.Text = lastWord; Testing
            string suggestedWord = "";
                suggestedWord += dataSet.SuggestNext(lastWord);
            Debug.Write("Suggested word: ", suggestedWord);
            IEnumerable<string> suggestedWords = dataSet.Next4Words(lastWord, 4);

            docCon += " " + suggestedWord;

           // foreach (string word in suggestedWords)
              //  {
                    //objPare.Range.Text = string.Concat(" ", suggestedWord);
                    //tbOutput.AppendText(string.Concat(" ", word));
                    //objPare.Range.Text.AppendText(string.Concat(" ", word));
               // }
            //objPare.Range.Text = textFromDoc;
            //objPare.Range.Text = docCon;
           // Debug.Write("testing2: ", lastWord);
           // Debug.Write("testing Doc: ", docCon);

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
        private void OnDataSetLoaded()
        {
            labelTotalWords.Visible = true;
            labelUniqueWords.Visible = true;
            IsDatasetDirty = false;
            UpdateLabels();
        }
        private void UpdateLabels()
        {
            labelTotalWords.Label = string.Format("{0} Total Words", dataSet.TotalSampleSize);
            labelUniqueWords.Label = string.Format("{0} Unique Words", dataSet.UniqueWordCount);
        }
        private void OpenDataSet()
        {
            //if (AskIfSaveFirst())
            //{
                string selectedFile = "C:\\Users\\kuro0\\source\\repos\\SeniorDesign\\SeniorDesign\\Texts\\Dictionary.txt";
                    dataSet = TrainedDataSet.DeserializeFromXml(selectedFile);
                    OnDataSetLoaded();

            //}
        }

        /*
    private void TrainDataSet()
    {
        string selectedFile = ShowFileDialog(openFileDialog);
        if (!string.IsNullOrWhiteSpace(selectedFile) && File.Exists(selectedFile))
        {
            dataSet.Train(new FileInfo(selectedFile));

            IsDatasetDirty = true;
            UpdateLabels();
        }
    }

    private string ShowFileDialog(FileDialog dialog)
    {
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            return dialog.FileName;
        }
        else
        {
            return string.Empty;
        }
    }

    private void NewDataSet()
    {
        if (AskIfSaveFirst())
        {
            dataSet = new TrainedDataSet();
            OnDataSetLoaded();
        }
    }

    */
    }

    }


