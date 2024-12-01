using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;

namespace FlightApp
{
    public partial class FormFlightApp : Form
    {
        /*
         * Naam: Josse Denis
         * Klas: 6ADB
         */
        private List<string[]> _filteredFlightsList = new List<string[]>();
        private List<string[]> _flightsList = new List<string[]>();
        public FormFlightApp()
        {
            InitializeComponent();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
            // vergeet niet via " add reference" het microsoft excel objext toe te voegen
            // declatatie
            Microsoft.Office.Interop.Excel.Application xlToep = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWerkmap;
            Microsoft.Office.Interop.Excel.Worksheet xlWerkblad;
            int rij = 2;
            // declaratie
            string departure_id, departure, arrivel_id, arrivel_city, max_capacity, actual_capacity, date_of_flight, average_cost, type_of_flight;

            

            // instellen van dialoogvenster openen
            OpenFileDialog dlgOpen = new OpenFileDialog();
            // Eigenschappen instellen
            dlgOpen.Title = "Openen";
            dlgOpen.FileName = "";
            dlgOpen.DefaultExt = ".xlsx";
            dlgOpen.InitialDirectory = Application.StartupPath;
            dlgOpen.Filter = "exelBestand (.xlsx) |*.xlsx|Alle Bestanden (*.*)|*.*";

            // Dialoogvenster tonen en de keuzen opvangen
            DialogResult resultaat = dlgOpen.ShowDialog();
            // openen van de werkmap
            xlWerkmap = xlToep.Workbooks.Open(dlgOpen.FileName);
            // openen werkblad
            xlWerkblad = xlWerkmap.ActiveSheet;

            // kijken of de gebuiker "openen" (ok) geklikt heeft.
            if (resultaat == DialogResult.OK)
            {
                

               

                // keuzelijst leegmaken
                listBoxFlights.Items.Clear();

                while (xlWerkblad.Cells[rij,1].Value != null)
                {
                    string[] flight = new string[9];

                    flight[0] = xlWerkblad.Cells[rij, 1].Value.ToString();
                    flight[1] = xlWerkblad.Cells[rij, 2].Value.ToString();
                    flight[2] = xlWerkblad.Cells[rij, 3].Value.ToString();
                    flight[3] = xlWerkblad.Cells[rij, 4].Value.ToString();
                    flight[4] = xlWerkblad.Cells[rij, 5].Value.ToString();
                    flight[5] = xlWerkblad.Cells[rij, 6].Value.ToString();
                    flight[6] = xlWerkblad.Cells[rij, 7].Value.ToString();
                    flight[7] = xlWerkblad.Cells[rij, 8].Value.ToString();
                    flight[8] = xlWerkblad.Cells[rij, 9].Value.ToString();
                  
                    _flightsList.Add(flight);
                    rij++;


                }
                toonList(_flightsList);
                // excel afsluiten
                xlToep.Quit();
            }

            /*
             * Naam: Josse Denis
             * Klas: 6ADB
             */





        }

        private void toonList(List<string[]> flightsList)
        {
            foreach (string[] flight in flightsList)
            {
                string temp = "";
                temp = flight[0].PadRight(5);
                temp += flight[1].PadRight(20);
                temp += flight[2].PadRight(5);
                temp += flight[3].PadRight(15);
                temp += flight[4].PadRight(10);
                temp += flight[5].PadRight(5);
                temp += flight[6].Substring(0, 10).PadRight(10);
                temp += flight[7].PadRight(5);
                temp += flight[8].PadRight(5);
               
                
                listBoxFlights.Items.Add(temp);
            }
        }
    }
}
