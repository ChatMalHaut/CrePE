using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;


namespace CREPE
{
    public partial class ThisAddIn
    {
        Outlook.Explorer currentExplorer = null;

        string idducalendar { get; set; }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            currentExplorer = this.Application.ActiveExplorer();
            CREPE.rubanAddConge newRuban = new rubanAddConge();
            newRuban.PerformLayout();
            idducalendar = CreateCustomCalendar();

        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Remarque : Outlook ne déclenche plus cet événement. Si du code
            //    doit s'exécuter à la fermeture d'Outlook, voir http://go.microsoft.com/fwlink/?LinkId=506785
        }

        private string CreateCustomCalendar()
        {
            const string newCalendarName = "Congés collaborateurs";
            Outlook.Folder primaryCalendar = (Outlook.Folder)this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            string id = "";
            bool needFolder = true;

            foreach (Outlook.Folder personalCalendar in primaryCalendar.Folders)
            {
                if (personalCalendar.Name == newCalendarName)
                {
                    needFolder = false;
                    id = personalCalendar.EntryID;
                    break;
                }
            }
            if (needFolder)
            {
                Outlook.Folder personalCalendar = (Outlook.Folder)primaryCalendar.Folders.Add(newCalendarName, Outlook.OlDefaultFolders.olFolderCalendar);
                id = personalCalendar.EntryID;
            }
            return id;
        }

        public void creationDeConge()
        {
            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                int nombreMailValides = 0;
                int nombreMailNonValides = 0;
                int nombreObjetsNonMails = 0;
                foreach (Object selectedObject in this.Application.ActiveExplorer().Selection)
                {

                    
                    if (selectedObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem = selectedObject as Outlook.MailItem;

                        String syges = mailItem.SenderEmailAddress;
                        String Sujet = mailItem.Subject;
                        String Corps = mailItem.Body;

                        if (syges == "sygesweb@netapsys.fr" && Sujet.Contains("NETAP"))
                        //if (syges == "antonin.rousset@netapsys.fr" && Sujet.Contains("NETAP"))
                        {
                            nombreMailValides++;

                            // Âme sensible s'abstenir. Séparation des données suivant le "mail type" de demande de congés.

                            String Demandeur = Sujet.Split('-')[2];
                            String TypeDeConge = Corps.Split('-')[1].Split(':')[1];
                            String DateDebut = Corps.Split('-')[2].Split('u')[1].Split('a')[0];
                            String DateFin = Corps.Split('-')[2].Split('u')[2].Split(null)[1];

                            // Retrait des espaces

                            DateFin = DateFin.Trim();
                            DateDebut = DateDebut.Trim();
                            TypeDeConge = TypeDeConge.Trim();

                            createConge(Demandeur, DateDebut, DateFin, TypeDeConge);

                        }

                        else
                        {
                            nombreMailNonValides++;
                        }


                    }
                    else
                    {
                        nombreObjetsNonMails++;
                    }

                }
                MessageBox.Show(string.Format("Nombre de mails traités: {0} \n Nombre de mails non traités: {1} \n Nombre d'objets sélectionnés non valides: {2}", nombreMailValides, nombreMailNonValides, nombreObjetsNonMails));
            }
            else
            {
                MessageBox.Show("Rien n'est sélectionné \n Veuillez sélectionner un mail de demande de congé.");
            }
        }



        private void createConge(String nomDemandeur, String dateDebut, String dateFin, String typedeconge)
        {
            Outlook.AppointmentItem nouveauConge = this.Application.Session.GetFolderFromID(idducalendar).Items.Add(Outlook.OlItemType.olAppointmentItem);
            nouveauConge.Subject = "Congé : " + nomDemandeur;
            nouveauConge.Body = "Type du congé : " + typedeconge;
            nouveauConge.Start = DateTime.Parse(dateDebut);
            nouveauConge.End = DateTime.Parse(dateFin + " 12:00 PM");
            nouveauConge.AllDayEvent = true;
            nouveauConge.ReminderSet = false;
            
            // Recherche du congé (appointment) dans le calendrier
            Outlook.Folder folder = (Outlook.Folder)Globals.ThisAddIn.Application.Session.GetFolderFromID(idducalendar);

            string filter = "[Subject] =  'Congé : " + nomDemandeur + "'";
            object obj = folder.Items.Find(filter);

            if (obj != null)
            {
                //  le congé (appointment) existe déjà
                Outlook.AppointmentItem appointment = obj as Outlook.AppointmentItem;

                if (appointment.Start == nouveauConge.Start || appointment.End == nouveauConge.End)
                {
                    appointment.Delete();

                    nouveauConge.Save();
                }

                else
                {
                    // Nouveau congé

                    nouveauConge.Save();

                }
            }
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
