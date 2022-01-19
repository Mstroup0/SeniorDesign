using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Collections.Generic;
using WordPredictionLibrary.Core;

namespace SuggestWordLibrary
{
	public partial class MainForm : Form
	{
		#region Winform Event Handlers

		public MainForm()
		{
			InitializeComponent();
			dataSet = new TrainedDataSet();
			IsDatasetDirty = false;
		}

		private void btnNew_Click(object sender, EventArgs e)
		{
			NewDataSet();
		}

		private void btnSave_Click(object sender, EventArgs e)
		{
			SaveDataSet();
		}

		private void btnOpen_Click(object sender, EventArgs e)
		{
			OpenDataSet();
		}

		private void btnTrain_Click(object sender, EventArgs e)
		{
			TrainDataSet();
		}

		private void ShowInfo(string format, params object[] args)
		{
			MessageBox.Show(string.Format(format, args), "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void ShowError(string format, params object[] args)
		{
			MessageBox.Show(string.Format(format, args), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}

		private void tbOutput_KeyUp(object sender, KeyEventArgs e)
		{
			if (e.Control && e.KeyCode == Keys.A) // CTRL + A, Select all
			{
				tbOutput.SelectAll();
			}
		}

		#endregion

		bool IsDatasetDirty { get; set; }
		TrainedDataSet dataSet { get; set; }

		private void NewDataSet()
		{
			if (AskIfSaveFirst())
			{
				dataSet = new TrainedDataSet();
				OnDataSetLoaded();
			}
		}

		private void OpenDataSet()
		{
			if (AskIfSaveFirst())
			{
				string selectedFile = ShowFileDialog(openFileDialog);
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

		/// <summary>
		/// Asks to save the work. Returns false if Cancel was pressed.
		/// </summary>		
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

		private void OnDataSetLoaded()
		{
			btnTrain.Visible = true;
			labelTotalWords.Visible = true;
			labelUniqueWords.Visible = true;
			IsDatasetDirty = false;
			UpdateLabels();
		}

		private void SaveDataSet()
		{
			string selectedFile = ShowFileDialog(saveFileDialog);
			if (!string.IsNullOrWhiteSpace(selectedFile))
			{
				if (TrainedDataSet.SerializeToXml(dataSet, selectedFile))
				{
					IsDatasetDirty = false;
				}
			}
		}

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

		private void UpdateLabels()
		{
			labelTotalWords.Text = string.Format("{0} Total Words", dataSet.TotalSampleSize);
			labelUniqueWords.Text = string.Format("{0} Unique Words", dataSet.UniqueWordCount);
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

		private void btnSuggest_Click(object sender, EventArgs e)
		{
			string lastWord = extractLastWord(tbOutput.Text);
			string suggestedWord = dataSet.SuggestNext(lastWord);
			IEnumerable<string> suggestedWords = dataSet.Next4Words(lastWord, 4);
			tbOutput.AppendText(string.Concat(" ", suggestedWord));
			foreach (string word in suggestedWords)
			{
				tbOutput.AppendText(string.Concat(" ", word));
			}
		}

		private string extractLastWord(string text)
		{
			string input = text.TrimEnd(' ', '\t', '\n');
			if (!string.IsNullOrWhiteSpace(input))
			{
				int indexOfLastWord = input.LastIndexOf(' ');
				if (indexOfLastWord == -1)
				{
					indexOfLastWord = 0;
				}
				string lastWord = tbOutput.Text.Substring(indexOfLastWord).Trim();
				return lastWord;
			}
			return string.Empty;
		}
		
		private void btnDumpAll_Click(object sender, EventArgs e)
		{
			string selectedFile = ShowFileDialog(openFileDialog);
			if (!string.IsNullOrWhiteSpace(selectedFile))
			{
				File.WriteAllText(selectedFile, dataSet.GetEntireDictionaryString());
			}
		}
	}
}
