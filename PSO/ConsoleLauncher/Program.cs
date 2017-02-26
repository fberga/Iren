using Iren.PSO.Base;
using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Deployment.Application;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Mono.Options;
using System.Collections.Generic;
using System.Xml.Linq;

namespace Iren.PSO.ConsoleLauncher
{
    class Program
    {
        private static Excel.Application _xlApp;

        static void Main(string[] args)
        {
            // these variables will be set when the command line is parsed
            int idApplicazione = -1;
            bool accettaCambioData = false;
            bool rifiutaCambioData = false;
            bool aggiornaStruttura = false;
            bool aggiornaDati = false;
            bool eseguiAzioni = false;
            string listaAzioni = "";
            bool haEntita = false;
            string listaEntita = "";
            bool shouldShowHelp = false;

            // these are the available options, not that they set the variables
            OptionSet options = new OptionSet { 
                { "i=", "l'id dell'applicazione.", (int id) => 
                    {
                        idApplicazione = id;
                    }
                }, 
                { "a", "accetta automaticamente il cambio data", cd => accettaCambioData = cd != null }, 
                { "r", "rifiuta automaticamente il cambio data", rd => rifiutaCambioData = rd != null },
                { "s", "forza l'aggiornamento della struttura e dati", aggstr => aggiornaStruttura = aggstr != null},
                { "d", "forza l'aggiornamento dei dati ma non della struttura", aggdt => aggiornaDati = aggdt != null},
                { "l=", "la lista di azioni separate da ; (\"\" per tutte le azioni)", lista => 
                    {
                        eseguiAzioni = lista != null;
                        listaAzioni = lista;
                    }},
                { "e=", "la lista delle entita separate da ; (\"\" per tutte le entita)", lista => 
                    {
                        haEntita = lista != null;
                        listaEntita = lista;
                    }},
                { "h|help", "mostra questo help ed esce", h => shouldShowHelp = h != null }
            };


            List<string> extra;
            try
            {
                // parse the command line
                extra = options.Parse(args);

                if (idApplicazione == -1)
                    throw new OptionException("Manca l'ID dell'applicazione da avviare.", "-i");
            }
            catch (OptionException e)
            {
                // output some error message
                Console.Write("ConsoleLauncher: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `consolelauncher --help' for more information.");
                Console.ReadKey();
                return;
            }

            if (shouldShowHelp)
            {
                // show some app description message
                Console.WriteLine("Utilizzo: ConsoleLauncer.exe [OPTIONS]+");
                Console.WriteLine("Esegue un'applicazione della suite PSO dal terminale.");
                Console.WriteLine("L'IdApplicazione deve essere specificato.");
                Console.WriteLine();

                // output the options
                Console.WriteLine("Options:");
                options.WriteOptionDescriptions(Console.Out);
                return;
            }


            Excel.Workbooks wbs = null;
            try
            {
                wbs = _xlApp.Workbooks;
            }
            catch
            {
                _xlApp = new Excel.Application();
            }
            finally
            {
                if (wbs != null) Marshal.ReleaseComObject(wbs);
                wbs = null;
            }
            
            XDocument doc = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
               new XElement("AvvioAutomatico",
               new XElement("AccettaCambioData", accettaCambioData),
               new XElement("RifiutaCambioData", rifiutaCambioData),
               new XElement("AggiornaStruttura", aggiornaStruttura),
               new XElement("AggiornaDati", aggiornaDati)));

            if (eseguiAzioni)
            {
                doc.Element("AvvioAutomatico").Add(
                    new XElement("ListaAzioni", listaAzioni));
            }
            if (haEntita)
            {
                doc.Element("AvvioAutomatico").Add(
                    new XElement("ListaEntita", listaEntita));
            }

            doc.Save(@"C:\Emergenza\AvvioAutomatico.xml");

            //COMMENTATA PER SCOPI DI TEST
            Workbook.AvviaApplicazione(_xlApp, idApplicazione);
            
        }
    }
}
