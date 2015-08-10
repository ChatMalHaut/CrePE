using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {

        // Outlook.Inspectors inspectors;

        Outlook.Explorer currentExplorer = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Outlook.Inspector inspector = new Outlook.Inspector();
            currentExplorer = this.Application.ActiveExplorer();
            OutlookAddIn1.rubanAddConge newRuban = new rubanAddConge();
            //Outlook.Folder deletedFolder = (Outlook.Folder)this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
            //this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)
            //deletedFolder.Items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(DeletedItems_ItemAdd);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Remarque : Outlook ne déclenche plus cet événement. Si du code
            //    doit s'exécuter à la fermeture d'Outlook, voir http://go.microsoft.com/fwlink/?LinkId=506785
        }

        public void currentExplorer_Event()
        {
            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                foreach (Object selectedObject in this.Application.ActiveExplorer().Selection)
                {

                    //Object selectedObject = this.Application.ActiveExplorer().Selection[1];
                    if (selectedObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem = selectedObject as Outlook.MailItem;

                        String syges = mailItem.SenderEmailAddress;
                        String Sujet = mailItem.Subject;
                        String Corps = mailItem.Body;

                        if (syges == "sygesweb@netapsys.fr")
                        {

                            // Âme sensible s'abstenir. Séparation des données suivant le "mail type" de demande de congés.

                            String Demandeur = Sujet.Split('-')[2];
                            String TypeDeConge = Corps.Split('-')[1].Split(':')[1];
                            String DateDebut = Corps.Split('-')[2].Split('u')[1].Split('a')[0];
                            String DateFin = Corps.Split('-')[2].Split('u')[2].Split(null)[1];

                            // Retrait des espaces

                            DateFin = DateFin.Trim();
                            DateDebut = DateDebut.Trim();
                            TypeDeConge = TypeDeConge.Trim();

                            // Concaténation des données

                            //String parsed = demandeur; ";" + TypeDeConge + ";" + DateDebut + ";" + DateFin

                            //MessageBox.Show(parsed);

                            createConge(Demandeur, DateDebut, DateFin, TypeDeConge);



                        }
                        else
                        {
                            MessageBox.Show("Ceci n'est pas un mail de congé");
                        }


                    }
                }
            }
        }
        private void createConge(String nomDemandeur, String dateDebut, String dateFin, String typedeconge)
        {
            //String nomDemandeur = detailsConge.Split(';')[0];
            //String typedeconge = detailsConge.Split(';')[1];
            //String dateDebut = detailsConge.Split(';')[2];
            //String dateFin = detailsConge.Split(';')[3];
            Outlook.AppointmentItem nouveauConge = (Outlook.AppointmentItem)this.Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
            nouveauConge.Subject = "Congé : " + nomDemandeur;
            nouveauConge.Body = "Type du congé : " + typedeconge;
            nouveauConge.Start = DateTime.Parse(dateDebut);
            nouveauConge.End = DateTime.Parse(dateFin + " 12:00 PM");
            nouveauConge.AllDayEvent = true;
            nouveauConge.ReminderSet = false;
            nouveauConge.Save();
            //nouveauConge.Display(true);
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

        #endregion
    }
}
