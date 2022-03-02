using WordPredictionLibrary.Core;

namespace SeniorDesign
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
           // dataSet = new TrainedDataSet();
           // IsDatasetDirty = false;
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
            this.StartStop = this.Factory.CreateRibbonToggleButton();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.labelTotalWords = this.Factory.CreateRibbonLabel();
            this.labelUniqueWords = this.Factory.CreateRibbonLabel();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.group7 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            group1 = this.Factory.CreateRibbonGroup();
            group1.SuspendLayout();
            this.tab1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.group6.SuspendLayout();
            this.group7.SuspendLayout();
            this.SuspendLayout();
            // 
            // group1
            // 
            group1.Items.Add(this.StartStop);
            group1.Label = "group1";
            group1.Name = "group1";
            // 
            // StartStop
            // 
            this.StartStop.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.StartStop.Label = "Start";
            this.StartStop.Name = "StartStop";
            this.StartStop.ShowImage = true;
            this.StartStop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StartStop_Click);
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Groups.Add(this.group7);
            this.tab1.Label = "Senior Design";
            this.tab1.Name = "tab1";
            this.tab1.Tag = "";
            // 
            // group3
            // 
            this.group3.Items.Add(this.labelTotalWords);
            this.group3.Items.Add(this.labelUniqueWords);
            this.group3.Label = "Word Totals";
            this.group3.Name = "group3";
            // 
            // labelTotalWords
            // 
            this.labelTotalWords.Label = "{0} Total Words";
            this.labelTotalWords.Name = "labelTotalWords";
            this.labelTotalWords.Visible = false;
            // 
            // labelUniqueWords
            // 
            this.labelUniqueWords.Label = "{0} Unique Words";
            this.labelUniqueWords.Name = "labelUniqueWords";
            this.labelUniqueWords.Visible = false;
            // 
            // group4
            // 
            this.group4.Items.Add(this.b1Word);
            this.group4.Label = "Word 1";
            this.group4.Name = "group4";
            // 
            // group5
            // 
            this.group5.Items.Add(this.button3);
            this.group5.Label = "Word 2";
            this.group5.Name = "group5";
            // 
            // group6
            // 
            this.group6.Items.Add(this.button4);
            this.group6.Label = "Word 3";
            this.group6.Name = "group6";
            // 
            // group7
            // 
            this.group7.Items.Add(this.button5);
            this.group7.Label = "Word 4";
            this.group7.Name = "group7";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "button1";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Label = "button3";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click_1);
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Label = "button4";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.b4Word.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.b4Word.Label = "Word 4";
            this.b4Word.Name = "b4Word";
            this.b4Word.ShowImage = true;
            this.b4Word.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.b4Word_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            group1.ResumeLayout(false);
            group1.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group7.ResumeLayout(false);
            this.group7.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton StartStop;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelTotalWords;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelUniqueWords;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton b1Word;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton b2Word;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton b3Word;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton b4Word;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
