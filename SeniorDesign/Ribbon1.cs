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
        // start/stop button
        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {


        }
        private void button2_Click(object sender, RibbonControlEventArgs e)
       {
            dataSet = new TrainedDataSet();
            IsDatasetDirty = false;
            OpenDataSet();
            Globals.ThisAddIn.Suggest(); 
            
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
            if (AskIfSaveFirst())
            {
                //string selectedFile = ShowFileDialog(openFileDialog);
                string selectedFile = "C:\\Users\\kuro0\\source\\repos\\SeniorDesign\\SeniorDesign\\Texts\\Dictionary.txt";
                //Debug.WriteLine("file " + selectedFile);
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

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            string suggestion1 = Globals.ThisAddIn.arrayWords(0);
           // button1.Label = suggestion1;

        }

        private void button3_Click_1(object sender, RibbonControlEventArgs e)
        {
            string suggestion2 = Globals.ThisAddIn.arrayWords(1);
           // button3.Label = suggestion2;
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            string suggestion3 = Globals.ThisAddIn.arrayWords(2);
           // button4.Label = suggestion3;

        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            string suggestion4 = Globals.ThisAddIn.arrayWords(3);
           // button4.Label = suggestion2;
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


