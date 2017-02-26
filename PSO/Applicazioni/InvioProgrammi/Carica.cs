using Iren.PSO.Base;
using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Funzioni di caricamento personalizzato. Una volta caricati i dati, scrive l'informazione anche nei fogli di export.
    /// </summary>
    public class Carica : Base.Carica
    {
        DefinedNames _definedNamesSheetMercato = new DefinedNames(Workbook.Mercato);
        Excel.Worksheet _wsMercato;

        public Carica() 
            : base() 
        {
            _wsMercato = Workbook.Sheets[Workbook.Mercato];
        }

        public override bool AzioneInformazione(object siglaEntita, object siglaAzione, object azionePadre, DateTime giorno, string[] mercati, object parametro = null)
        {
            bool isGenera = azionePadre.Equals("GENERA");
            if (isGenera)
                azionePadre = "CARICA";

            bool o = base.AzioneInformazione(siglaEntita, siglaAzione, azionePadre, giorno, mercati, parametro);

            if(isGenera)
                azionePadre = "GENERA";

            //non ho fatto nulla, la connessione non si apre e l'azione padre è CARICA... rientro nel caso del caricamento da XML
            if (o == false && !DataBase.OpenConnection() && azionePadre.Equals("CARICA")) 
            {
                //tipo file da caricare
                string tf = siglaAzione.ToString();
                DataTable azioneInformazione = CaricaXML("pathExportFile" + tf, "formatoNomeFile" + tf, siglaEntita, tf) ?? new DataTable();

                if (azioneInformazione.Rows.Count == 0)
                {
                    DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno, false);
                    return false;
                }
                else
                {
                    string sheet = DefinedNames.GetSheetName(siglaEntita);
                    DefinedNames definedNames = new DefinedNames(sheet);

                    ScriviInformazione(siglaEntita, azioneInformazione.DefaultView, definedNames);
                    DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno);
                }
            }

            DataBase.CloseConnection();
            string name = DefinedNames.GetSheetName(siglaEntita);
            Sheet s = new Sheet(Workbook.Sheets[name]);
            s.AggiornaColori();

            return o;
        }

        private DataTable CaricaXML(string pathCfg, string nameFormatCfg, object siglaEntita, string tipoFile)
        {
            string path = Esporta.PreparePath(Workbook.GetUsrConfigElement(pathCfg));
            var name = Workbook.GetUsrConfigElement(nameFormatCfg);

            if(!Directory.Exists(path))
            {
                //TODO segnalare directory non accessibile
                return null;
            }

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            string codiceRUP = categoriaEntita[0]["CodiceRUP"].ToString();


            string nomeFile = Esporta.PrepareName(name.Value, codiceRUP) + "*.xml";


            string[] files = Directory.GetFiles(path, nomeFile, SearchOption.TopDirectoryOnly);

            if (files.Length > 0)
            {
                foreach (string file in files)
                {
                    if (tipoFile == "US")
                        return LeggiUS(file, siglaEntita.ToString(), codiceRUP);
                    else if (tipoFile == "FMS")
                        return LeggiFMS(file, siglaEntita.ToString(), codiceRUP);
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Nessun file " + tipoFile + " trovato per l'UP " + codiceRUP, Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }

            return null;
        }

        private DataTable LeggiUS(string file, string siglaEntita, string codiceRUP)
        {
            XDocument fileUS = XDocument.Load(file);
            XNamespace ns = fileUS.Root.Name.Namespace;//"urn:XML-PIPE";

            DataTable oDT = new DataTable()
            {
                Columns =
                {
                    {"SiglaEntita", typeof(string)},
                    {"SiglaInformazione", typeof(string)},
                    {"Data", typeof(string)},
                    {"Valore", typeof(double)},
                    {"BackColor", typeof(int)},
                    {"ForeColor", typeof(int)},
                    {"Commento", typeof(double)}
                }
            };

            var PIPTransactions = fileUS.Element(ns + "PIPEDocument").Elements(ns + "PIPTransaction");

            foreach (var PIPTransaction in PIPTransactions)
            {
                var UnitSchedule = PIPTransaction.Element(ns + "UnitSchedule");
                if (UnitSchedule.Attribute("Cummulative").Value.ToUpper() != "NO")
                {
                    if(UnitSchedule.Element(ns + "UnitReferenceNumber").Value == codiceRUP)
                    {
                        string date = "";
                        foreach (var ele in UnitSchedule.Elements())
                        {
                            if (ele.Name == ns + "Date")
                            {
                                date = ele.Value;
                            }
                            else if(ele.Name == ns + "Quantity")
                            {
                                DataRow r = oDT.NewRow();
                                r["SiglaEntita"] = siglaEntita;
                                r["SiglaInformazione"] = "PROGRAMMA_" + Workbook.Mercato;
                                r["Data"] = date + int.Parse(ele.Attribute("Hour").Value).ToString("00");
                                r["Valore"] = double.Parse(ele.Value, CultureInfo.InstalledUICulture);
                                oDT.Rows.Add(r);
                            }
                        }
                    }
                }
            }
            return oDT;
        }
        private DataTable LeggiFMS(string file, string siglaEntita, string codiceRUP)
        {
            XDocument fileUS = XDocument.Load(file);
            XNamespace ns = fileUS.Root.Name.Namespace;

            DataTable oDT = new DataTable()
            {
                Columns =
                {
                    {"SiglaEntita", typeof(string)},
                    {"SiglaInformazione", typeof(string)},
                    {"Data", typeof(string)},
                    {"Valore", typeof(double)},
                    {"BackColor", typeof(int)},
                    {"ForeColor", typeof(int)},
                    {"Commento", typeof(double)}
                }
            };

            var PIPTransactions = fileUS.Element(ns + "PIPEDocument").Elements(ns + "PIPTransaction");

            foreach (var PIPTransaction in PIPTransactions)
            {
                var FifteenMinuteSchedule = PIPTransaction.Element(ns + "FifteenMinuteSchedule");
                if (FifteenMinuteSchedule.Element(ns + "UnitReferenceNumber").Value == codiceRUP)
                {
                    string date = FifteenMinuteSchedule.Element(ns + "Date").Value;
                    var HourDetails = FifteenMinuteSchedule.Elements(ns + "HourDetail");

                    foreach (var ele in HourDetails)
                    {
                        for (int i = 1; i < 5; i++)
                        {
                            DataRow r = oDT.NewRow();
                            r["SiglaEntita"] = siglaEntita;
                            r["SiglaInformazione"] = "PROGRAMMAQ" + i + "_" + Workbook.Mercato;
                            r["Data"] = date + int.Parse(ele.Element(ns + "Hour").Value).ToString("00");
                            r["Valore"] = double.Parse(ele.Elements(ns + "Quantity").Where(e => e.Attribute("QuarterInterval").Value == i.ToString()).First().Value, CultureInfo.InstalledUICulture);
                            oDT.Rows.Add(r);
                        }
                    }
                }
            }
            return oDT;
        }

        protected override void ScriviCella(Excel.Worksheet ws, DefinedNames definedNames, object siglaEntita, DataRowView info, string suffissoData, string suffissoOra, object risultato, bool saveToDB, bool fromCarica)
        {
            base.ScriviCella(ws, definedNames, siglaEntita, info, suffissoData, suffissoOra, risultato, saveToDB, fromCarica);
            
            //se l'informazione è visibile la devo scrivere anche nei fogli dei mercati
            DataView informazioni = new DataView(Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE]);
            informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' OR SiglaEntitaRif = '" + siglaEntita + "' AND SiglaInformazione = '" + info["SiglaInformazione"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            bool visible = false;
            foreach (DataRowView r in informazioni)
                if (r["Visibile"].Equals("1"))
                    visible = true;

            if (visible)
            {
                DataTable entita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA];

                var rif =
                    (from r in entita.AsEnumerable()
                     where r["IdApplicazione"].Equals(Workbook.IdApplicazione) && r["SiglaEntita"].Equals(siglaEntita)
                     select new { SiglaEntita = r["Gerarchia"] is DBNull ? r["SiglaEntita"] : r["Gerarchia"], Riferimento = r["Riferimento"] }).First();

                string quarter = Regex.Match(info["SiglaInformazione"].ToString(), @"Q\d").Value;
                quarter = quarter == "" ? "Q1" : quarter;

                Range rngMercato = new Range(_definedNamesSheetMercato.GetRowByName(rif.SiglaEntita, "UM", "T") + 2, _definedNamesSheetMercato.GetColFromName("RIF" + rif.Riferimento, "PROGRAMMA" + quarter));
                rngMercato.StartRow += (Date.GetOraFromSuffissoOra(suffissoOra) - 1);

                _wsMercato.Range[rngMercato.ToString()].Value = risultato;
            }
        }
    }
}
