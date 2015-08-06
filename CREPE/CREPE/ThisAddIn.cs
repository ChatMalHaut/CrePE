using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace CREPE
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;

        Outlook.Folder newFolder;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Remarque : Outlook ne déclenche plus cet événement. Si du code
            //    doit s'exécuter à la fermeture d'Outlook, voir http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region Code généré par VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.SenderEmailAddress == "nassim.belkaid@netapsys.fr")
                {
                    mailItem.Move(newFolder);
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }

            }
        }

        #endregion
    }
}
