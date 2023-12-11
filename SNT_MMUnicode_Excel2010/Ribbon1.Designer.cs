namespace SNT_MMUnicode_Excel2010
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
            this.SNT_MMUnicode = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.ZGToUni = this.Factory.CreateRibbonButton();
            this.U_Zaw = this.Factory.CreateRibbonButton();
            this.btnSavePdf = this.Factory.CreateRibbonButton();
            this.btnXPS = this.Factory.CreateRibbonButton();
            this.SNT_MMUnicode.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            // 
            // SNT_MMUnicode
            // 
            this.SNT_MMUnicode.Groups.Add(this.group1);
            this.SNT_MMUnicode.Groups.Add(this.group2);
            this.SNT_MMUnicode.Label = "SNT_MMUnicode";
            this.SNT_MMUnicode.Name = "SNT_MMUnicode";
            // 
            // group1
            // 
            this.group1.Items.Add(this.ZGToUni);
            this.group1.Items.Add(this.U_Zaw);
            this.group1.Label = "Convert";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnSavePdf);
            this.group2.Items.Add(this.btnXPS);
            this.group2.Label = "Save";
            this.group2.Name = "group2";
            // 
            // ZGToUni
            // 
            this.ZGToUni.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ZGToUni.Image = global::SNT_MMUnicode_Excel2010.Properties.Resources.Unicode;
            this.ZGToUni.Label = "ZGToUni";
            this.ZGToUni.Name = "ZGToUni";
            this.ZGToUni.ShowImage = true;
            this.ZGToUni.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ZGToUni_Click);
            // 
            // U_Zaw
            // 
            this.U_Zaw.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.U_Zaw.Image = global::SNT_MMUnicode_Excel2010.Properties.Resources.Zg;
            this.U_Zaw.Label = "Uni To ZG";
            this.U_Zaw.Name = "U_Zaw";
            this.U_Zaw.ShowImage = true;
            this.U_Zaw.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.U_Zaw_Click);
            // 
            // btnSavePdf
            // 
            this.btnSavePdf.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSavePdf.Image = global::SNT_MMUnicode_Excel2010.Properties.Resources.pdf1;
            this.btnSavePdf.Label = "Save as Pdf";
            this.btnSavePdf.Name = "btnSavePdf";
            this.btnSavePdf.ShowImage = true;
            this.btnSavePdf.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSavePdf_Click);
            // 
            // btnXPS
            // 
            this.btnXPS.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnXPS.Image = global::SNT_MMUnicode_Excel2010.Properties.Resources.XPS;
            this.btnXPS.Label = "Save as XPS";
            this.btnXPS.Name = "btnXPS";
            this.btnXPS.ShowImage = true;
            this.btnXPS.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnXPS_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.SNT_MMUnicode);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.SNT_MMUnicode.ResumeLayout(false);
            this.SNT_MMUnicode.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab SNT_MMUnicode;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ZGToUni;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSavePdf;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnXPS;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton U_Zaw;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
