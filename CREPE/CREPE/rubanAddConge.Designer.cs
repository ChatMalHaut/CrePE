namespace CREPE
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.AddCongeGroup = this.Factory.CreateRibbonGroup();
            this.AjoutConge = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.AddCongeGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.AddCongeGroup);
            this.tab1.Label = "Congés";
            this.tab1.Name = "tab1";
            // 
            // AddCongeGroup
            // 
            this.AddCongeGroup.Items.Add(this.AjoutConge);
            this.AddCongeGroup.Name = "AddCongeGroup";
            // 
            // AjoutConge
            // 
            this.AjoutConge.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AjoutConge.Description = "Ajoute les informations du congé contenues dans le mail";
            this.AjoutConge.Image = global::CREPE.Properties.Resources.valise;
            this.AjoutConge.Label = "Ajout des congés";
            this.AjoutConge.Name = "AjoutConge";
            this.AjoutConge.ScreenTip = "Sélectionner une demande de congé et cliquer pour l\'ajouter au calendrier";
            this.AjoutConge.ShowImage = true;
            this.AjoutConge.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AjoutConge_Click);
            // 
            // rubanAddConge
            // 
            this.Name = "rubanAddConge";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.rubanAddConge_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.AddCongeGroup.ResumeLayout(false);
            this.AddCongeGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AddCongeGroup;
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
