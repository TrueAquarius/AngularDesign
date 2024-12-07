namespace PowerPoint.AddIn
{
    partial class AngularDesignRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AngularDesignRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen, andernfalls "false".</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupAngleSetup = this.Factory.CreateRibbonGroup();
            this.editBoxAngle = this.Factory.CreateRibbonEditBox();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonApply = this.Factory.CreateRibbonButton();
            this.groupAlignment = this.Factory.CreateRibbonGroup();
            this.buttonAlignLeft = this.Factory.CreateRibbonButton();
            this.buttonAlignRight = this.Factory.CreateRibbonButton();
            this.buttonAlignCenter = this.Factory.CreateRibbonButton();
            this.buttonStretch = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupAngleSetup.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupAlignment.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupAngleSetup);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.groupAlignment);
            this.tab1.Label = "Angular Design";
            this.tab1.Name = "tab1";
            // 
            // groupAngleSetup
            // 
            this.groupAngleSetup.Items.Add(this.editBoxAngle);
            this.groupAngleSetup.Label = "Angle Setup";
            this.groupAngleSetup.Name = "groupAngleSetup";
            // 
            // editBoxAngle
            // 
            this.editBoxAngle.Label = "Angle";
            this.editBoxAngle.Name = "editBoxAngle";
            this.editBoxAngle.Text = "10";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonApply);
            this.group1.Label = "Shape";
            this.group1.Name = "group1";
            // 
            // buttonApply
            // 
            this.buttonApply.Description = "Apply the given angle to the slope of the shapes";
            this.buttonApply.Label = "Apply angle to shape";
            this.buttonApply.Name = "buttonApply";
            this.buttonApply.SuperTip = "Apply the given angle to the slope of the shapes";
            this.buttonApply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonApply_Click);
            // 
            // groupAlignment
            // 
            this.groupAlignment.Items.Add(this.buttonAlignLeft);
            this.groupAlignment.Items.Add(this.buttonAlignRight);
            this.groupAlignment.Items.Add(this.buttonAlignCenter);
            this.groupAlignment.Items.Add(this.buttonStretch);
            this.groupAlignment.Label = "Alignment";
            this.groupAlignment.Name = "groupAlignment";
            // 
            // buttonAlignLeft
            // 
            this.buttonAlignLeft.Description = "Align left along a slope with given angle";
            this.buttonAlignLeft.Label = "Left";
            this.buttonAlignLeft.Name = "buttonAlignLeft";
            this.buttonAlignLeft.SuperTip = "Align left along a slope with given angle";
            this.buttonAlignLeft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAlignLeft_Click);
            // 
            // buttonAlignRight
            // 
            this.buttonAlignRight.Description = "Align right along a slope with given angle";
            this.buttonAlignRight.Label = "Right";
            this.buttonAlignRight.Name = "buttonAlignRight";
            this.buttonAlignRight.SuperTip = "Align right along a slope with given angle";
            this.buttonAlignRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAlignRight_Click);
            // 
            // buttonAlignCenter
            // 
            this.buttonAlignCenter.Description = "Align centered along a slope with given angle";
            this.buttonAlignCenter.Label = "Center";
            this.buttonAlignCenter.Name = "buttonAlignCenter";
            this.buttonAlignCenter.SuperTip = "Align centered along a slope with given angle";
            this.buttonAlignCenter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAlignCenter_Click);
            // 
            // buttonStretch
            // 
            this.buttonStretch.Description = "Stretch all selected shapes to equal length.";
            this.buttonStretch.Label = "Stretch";
            this.buttonStretch.Name = "buttonStretch";
            this.buttonStretch.SuperTip = "Stretch all selected shapes to equal length.";
            this.buttonStretch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonStretch_Click);
            // 
            // AngularDesignRibbon
            // 
            this.Name = "AngularDesignRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AngularDesignRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupAngleSetup.ResumeLayout(false);
            this.groupAngleSetup.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.groupAlignment.ResumeLayout(false);
            this.groupAlignment.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAngleSetup;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxAngle;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonApply;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAlignment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAlignLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAlignRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAlignCenter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStretch;
    }

    partial class ThisRibbonCollection
    {
        internal AngularDesignRibbon AngularDesignRibbon
        {
            get { return this.GetRibbon<AngularDesignRibbon>(); }
        }
    }
}
