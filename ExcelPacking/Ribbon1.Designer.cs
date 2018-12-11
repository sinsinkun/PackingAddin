namespace ExcelPacking
{
    partial class ArrayRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ArrayRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ArrayRibbon));
            this.Tab_Addons = this.Factory.CreateRibbonTab();
            this.Group_Packing = this.Factory.CreateRibbonGroup();
            this.Btn_Pack = this.Factory.CreateRibbonButton();
            this.Tab_Addons.SuspendLayout();
            this.Group_Packing.SuspendLayout();
            this.SuspendLayout();
            // 
            // Tab_Addons
            // 
            this.Tab_Addons.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Tab_Addons.Groups.Add(this.Group_Packing);
            this.Tab_Addons.Label = "Array Add-ons";
            this.Tab_Addons.Name = "Tab_Addons";
            // 
            // Group_Packing
            // 
            this.Group_Packing.Items.Add(this.Btn_Pack);
            this.Group_Packing.Label = "Packing";
            this.Group_Packing.Name = "Group_Packing";
            // 
            // Btn_Pack
            // 
            this.Btn_Pack.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Btn_Pack.Image = ((System.Drawing.Image)(resources.GetObject("Btn_Pack.Image")));
            this.Btn_Pack.Label = "Pack";
            this.Btn_Pack.Name = "Btn_Pack";
            this.Btn_Pack.ShowImage = true;
            this.Btn_Pack.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_Pack_Click);
            // 
            // ArrayRibbon
            // 
            this.Name = "ArrayRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.Tab_Addons);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.Tab_Addons.ResumeLayout(false);
            this.Tab_Addons.PerformLayout();
            this.Group_Packing.ResumeLayout(false);
            this.Group_Packing.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab_Addons;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group_Packing;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Btn_Pack;
    }

    partial class ThisRibbonCollection
    {
        internal ArrayRibbon Ribbon1
        {
            get { return this.GetRibbon<ArrayRibbon>(); }
        }
    }
}
