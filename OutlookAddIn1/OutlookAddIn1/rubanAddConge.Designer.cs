namespace OutlookAddIn1
{
    partial class rubanAddConge : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public rubanAddConge()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.AddConge = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.AjoutConge = this.Factory.CreateRibbonButton();
            this.AddConge.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // AddConge
            // 
            this.AddConge.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.AddConge.Groups.Add(this.group1);
            this.AddConge.Label = "Ajout de Congé";
            this.AddConge.Name = "AddConge";
            // 
            // group1
            // 
            this.group1.Items.Add(this.AjoutConge);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // AjoutConge
            // 
            this.AjoutConge.Label = "Ajout du congé";
            this.AjoutConge.Name = "AjoutConge";
            this.AjoutConge.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AjoutConge_Click);
            // 
            // rubanAddConge
            // 
            this.Name = "rubanAddConge";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.AddConge);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.rubanAddConge_Load);
            this.AddConge.ResumeLayout(false);
            this.AddConge.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab AddConge;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AjoutConge;
    }

    partial class ThisRibbonCollection
    {
        internal rubanAddConge rubanAddConge
        {
            get { return this.GetRibbon<rubanAddConge>(); }
        }
    }
}
