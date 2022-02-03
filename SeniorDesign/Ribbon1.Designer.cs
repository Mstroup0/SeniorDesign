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
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.labelTotalWords = this.Factory.CreateRibbonLabel();
            this.labelUniqueWords = this.Factory.CreateRibbonLabel();
            this.tab2 = this.Factory.CreateRibbonTab();
            group1 = this.Factory.CreateRibbonGroup();
            group1.SuspendLayout();
            this.tab1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.tab2.SuspendLayout();
            this.SuspendLayout();
            // 
            // group1
            // 
            group1.Items.Add(this.toggleButton1);
            group1.Label = "group1";
            group1.Name = "group1";
            // 
            // toggleButton1
            // 
            this.toggleButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButton1.Label = "toggleButton1";
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.ShowImage = true;
            this.toggleButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click);
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Senior Design";
            this.tab1.Name = "tab1";
            this.tab1.Tag = "";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button2);
            this.group2.Label = "Words";
            this.group2.Name = "group2";
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Label = "Test";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.labelTotalWords);
            this.group3.Items.Add(this.labelUniqueWords);
            this.group3.Label = "group3";
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
            // tab2
            // 
            this.tab2.Label = "Senoir Design";
            this.tab2.Name = "tab2";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            group1.ResumeLayout(false);
            group1.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelTotalWords;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelUniqueWords;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
