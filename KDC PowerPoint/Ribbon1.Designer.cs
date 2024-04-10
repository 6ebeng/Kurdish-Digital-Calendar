namespace KDC_PowerPoint
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.splitButton3 = this.Factory.CreateRibbonSplitButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.checkBox7 = this.Factory.CreateRibbonCheckBox();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.dropDown2 = this.Factory.CreateRibbonDropDown();
            this.dropDown3 = this.Factory.CreateRibbonDropDown();
            this.dropDown4 = this.Factory.CreateRibbonDropDown();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.splitButton2 = this.Factory.CreateRibbonSplitButton();
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.checkBox2 = this.Factory.CreateRibbonCheckBox();
            this.checkBox3 = this.Factory.CreateRibbonCheckBox();
            this.checkBox4 = this.Factory.CreateRibbonCheckBox();
            this.checkBox5 = this.Factory.CreateRibbonCheckBox();
            this.checkBox6 = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.group4.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "KD Calendar";
            this.tab1.Name = "tab1";
            // 
            // group4
            // 
            this.group4.Items.Add(this.splitButton3);
            this.group4.Label = " ";
            this.group4.Name = "group4";
            // 
            // splitButton3
            // 
            this.splitButton3.ButtonEnabled = false;
            this.splitButton3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton3.Image = ((System.Drawing.Image)(resources.GetObject("splitButton3.Image")));
            this.splitButton3.Items.Add(this.button1);
            this.splitButton3.Label = "Help";
            this.splitButton3.Name = "splitButton3";
            this.splitButton3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButton3_Click);
            // 
            // button1
            // 
            this.button1.Label = "Credits";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.dropDown1);
            this.group1.Items.Add(this.checkBox7);
            this.group1.Label = "Choose Dialect";
            this.group1.Name = "group1";
            // 
            // dropDown1
            // 
            this.dropDown1.Label = "dropDown1";
            this.dropDown1.Name = "dropDown1";
            this.dropDown1.ShowItemImage = false;
            this.dropDown1.ShowLabel = false;
            this.dropDown1.SizeString = "MY_MAX_LENGTH_STRING";
            this.dropDown1.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown1_SelectionChanged);
            // 
            // checkBox7
            // 
            this.checkBox7.Label = "Add suffix";
            this.checkBox7.Name = "checkBox7";
            this.checkBox7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox7_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.dropDown2);
            this.group3.Items.Add(this.dropDown3);
            this.group3.Items.Add(this.dropDown4);
            this.group3.Items.Add(this.separator1);
            this.group3.Items.Add(this.splitButton2);
            this.group3.Label = "Converter";
            this.group3.Name = "group3";
            // 
            // dropDown2
            // 
            this.dropDown2.Label = "Calendar";
            this.dropDown2.Name = "dropDown2";
            this.dropDown2.ShowItemImage = false;
            this.dropDown2.SizeString = "MY_MAX_LENGTH_STRING";
            this.dropDown2.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown2_SelectionChanged);
            // 
            // dropDown3
            // 
            this.dropDown3.Label = "from";
            this.dropDown3.Name = "dropDown3";
            this.dropDown3.ShowItemImage = false;
            this.dropDown3.SizeString = "MY_MAX_LENGTH_STRING";
            this.dropDown3.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown3_SelectionChanged);
            // 
            // dropDown4
            // 
            this.dropDown4.Label = "to";
            this.dropDown4.Name = "dropDown4";
            this.dropDown4.ShowItemImage = false;
            this.dropDown4.SizeString = "MY_MAX_LENGTH_STRING";
            this.dropDown4.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown4_SelectionChanged);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // splitButton2
            // 
            this.splitButton2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton2.Image = ((System.Drawing.Image)(resources.GetObject("splitButton2.Image")));
            this.splitButton2.Items.Add(this.toggleButton1);
            this.splitButton2.Label = "Convert";
            this.splitButton2.Name = "splitButton2";
            this.splitButton2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButton2_Click);
            // 
            // toggleButton1
            // 
            this.toggleButton1.Label = "Reverse";
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.ShowImage = true;
            this.toggleButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.splitButton1);
            this.group2.Label = "Insert date";
            this.group2.Name = "group2";
            // 
            // splitButton1
            // 
            this.splitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton1.Image = ((System.Drawing.Image)(resources.GetObject("splitButton1.Image")));
            this.splitButton1.Items.Add(this.checkBox1);
            this.splitButton1.Items.Add(this.checkBox2);
            this.splitButton1.Items.Add(this.checkBox3);
            this.splitButton1.Items.Add(this.checkBox4);
            this.splitButton1.Items.Add(this.checkBox5);
            this.splitButton1.Items.Add(this.checkBox6);
            this.splitButton1.Label = "Select";
            this.splitButton1.Name = "splitButton1";
            this.splitButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButton1_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "dddd, dd MMMM, yyyy";
            this.checkBox1.Name = "checkBox1";
            // 
            // checkBox2
            // 
            this.checkBox2.Label = "dddd, dd/MM/yyyy";
            this.checkBox2.Name = "checkBox2";
            // 
            // checkBox3
            // 
            this.checkBox3.Label = "dd MMMM, yyyy";
            this.checkBox3.Name = "checkBox3";
            // 
            // checkBox4
            // 
            this.checkBox4.Label = "dd/MM/yyyy";
            this.checkBox4.Name = "checkBox4";
            // 
            // checkBox5
            // 
            this.checkBox5.Label = "MM/dd/yyyy";
            this.checkBox5.Name = "checkBox5";
            // 
            // checkBox6
            // 
            this.checkBox6.Label = "yyyy/MM/dd";
            this.checkBox6.Name = "checkBox6";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox7;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown2;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown3;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown4;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox4;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox5;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox6;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
