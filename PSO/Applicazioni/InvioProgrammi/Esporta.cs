using Iren.PSO.Base;
using Iren.PSO.UserConfig;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Crea la mail con i dati di export da inviare agli impianti.
    /// </summary>
    class Esporta : AEsporta
    {
        DefinedNames _defNamesMercato = new DefinedNames(Workbook.Mercato);

        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif, string[] mercati)
        {
            DataView entitaAzione = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            if (entitaAzione.Count == 0)
                return false;

            switch (siglaAzione.ToString())
            {
                case "MAIL":
                    //carico i path di export
                    List<UserConfigElement> cfgPaths = new List<UserConfigElement>();

                    var cfgPath = Workbook.GetUsrConfigElement("pathExportFileFMS");
                    cfgPaths.Add(cfgPath);
                    cfgPath = Workbook.GetUsrConfigElement("pathExportFileXSD");
                    cfgPaths.Add(cfgPath);
                    cfgPath = Workbook.GetUsrConfigElement("pathExportFileRS");
                    cfgPaths.Add(cfgPath);

                    //verifico che siano tutti raggiungibili
                    foreach (var p in cfgPaths)
                    {
                        string path = PreparePath(p);
                        if (!Directory.Exists(path))
                        {
                            System.Windows.Forms.MessageBox.Show(p.Desc + " '" + path + "' non raggiungibile.", Simboli.NomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                            return false;
                        }
                    }

                    Globals.ThisWorkbook.Application.ScreenUpdating = false;

                    var oldActiveWindow = Globals.ThisWorkbook.Application.ActiveWindow;
                    Globals.ThisWorkbook.Worksheets[Workbook.Mercato].Activate();

                    Range rng = new Range(_defNamesMercato.GetRowByName(siglaEntita, "DATA"), 1, Date.GetOreGiorno(Workbook.DataAttiva) + 5, _defNamesMercato.GetLastCol());

                    bool result = InviaMail(Workbook.Mercato, siglaEntita, rng);
                 
                    oldActiveWindow.Activate();

                    Globals.ThisWorkbook.Application.ScreenUpdating = true;
                    return result;
            }

            return true;
        }
        protected bool CreaOutputXLS(Excel.Worksheet ws, string fileName, bool deleteOrco, Range rng)
        {
            bool hasVariations = false;

            Excel.Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Add();
            ws.Range[rng.ToString()].Copy();
            wb.Sheets[1].Range["A1"].PasteSpecial(Excel.XlPasteType.xlPasteAllUsingSourceTheme);

            //fisso la formattazione condizionale nel range copiato
            foreach (Range cell in rng.Cells)
            {
                //traslo la cella per il nuovo foglio
                Range tCell = new Range(cell);
                tCell.StartRow -= (rng.StartRow - 1);
                tCell.StartColumn -= (rng.StartColumn - 1);
                wb.Sheets[1].Range[tCell.ToString()].Interior.ColorIndex = ws.Range[cell.ToString()].DisplayFormat.Interior.ColorIndex;

                if (wb.Sheets[1].Range[tCell.ToString()].Interior.ColorIndex == Struct.COLORE_VARIAZIONE_NEGATIVA || wb.Sheets[1].Range[tCell.ToString()].Interior.ColorIndex == Struct.COLORE_VARIAZIONE_POSITIVA)
                    hasVariations = true;
            }
            //rimuovo la formattazione condizionale
            Excel.Range tab = wb.Sheets[1].UsedRange;
            tab.FormatConditions.Delete();

            if (deleteOrco)
            {
                //TODO CHECK se rimuovere...
                if(DateTime.Now < new DateTime(2014, 07, 01))
                    wb.Sheets[1].Columns[8].EntireColumn.Delete();

                wb.Sheets[1].Range["H3"].Value = "Programma indicativo ORCO";
                wb.Sheets[1].Range["H3"].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            }

            //salvo l'export e lo chiudo
            wb.Sheets[1].Range["A1"].Select();
            wb.SaveAs(fileName, Excel.XlFileFormat.xlExcel8);
            wb.Close();

            return hasVariations;
        }
        protected bool InviaMail(string nomeFoglio, object siglaEntita, Range rng) 
        {
            List<string> attachments = new List<string>();
            bool hasVariations = false;
            try
            {
                Excel.Worksheet ws = Globals.ThisWorkbook.Sheets[nomeFoglio];
                
                //inizializzo l'oggetto mail
                Outlook.Application outlook = GetOutlookInstance();
                Outlook._MailItem mail = outlook.CreateItem(Outlook.OlItemType.olMailItem);

                DataView entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_ALLEGATO_EXCEL' AND IdApplicazione = " + Workbook.IdApplicazione;
                if (entitaProprieta.Count > 0)
                {
                    //creo file Excel da allegare
                    string excelExport = Path.Combine(@"C:\Emergenza", Workbook.DataAttiva.ToString("yyyyMMdd") + "_" + entitaProprieta[0]["Valore"] + "_" + Workbook.Mercato + ".xls");
                    attachments.Add(excelExport);

                    hasVariations = CreaOutputXLS(ws, attachments.Last(), siglaEntita.Equals("CE_ORX"), rng);

                    //25/01/2017 ENH: path export da file di configurazione
                    var cfgPath = Workbook.GetUsrConfigElement("pathExportFileRosone");

                    //18/01/2017 FIX: inizializzazione fuori dall'if: provocava errore invio mail
                    string pathExport = PreparePath(cfgPath);
                    string fileNameXML = null;
                    string fileNameCSV = null;
                    string pathXmlOrcoExportFull = null;
                    if (siglaEntita.Equals("CE_ORX"))
                    {
                        if (!Directory.Exists(pathExport))
                            Directory.CreateDirectory(pathExport);

                        fileNameXML = fileNameCSV = "FMS_UP_ORCO_1_" + Workbook.Mercato + "D_" + Workbook.DataAttiva.ToString("yyyyMMdd");

                        fileNameXML += ".xml.OEIESRD.out.xml";
                        fileNameCSV += ".csv.OEIESRD.out.csv";
                        pathXmlOrcoExportFull = pathExport + "\\" + fileNameXML;

                        bool xmlCreated = CreaOutputXML(siglaEntita, pathExport, fileNameXML, Workbook.DataAttiva);
                        bool csvCreated = CreaOutputCSV(siglaEntita, pathExport, fileNameCSV, Workbook.DataAttiva);

                        attachments.Add(Path.Combine(pathXmlOrcoExportFull));
                    }

                    DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                    categoriaEntita.RowFilter = "Gerarchia = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;

                    if(categoriaEntita.Count == 0)
                        categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;

                    bool interrupt = false;

                    foreach (DataRowView entita in categoriaEntita)
                    {
                        entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_ALLEGATO_FMS' AND IdApplicazione = " + Workbook.IdApplicazione;
                        if (entitaProprieta.Count > 0)
                        {
                            //cerco i file XML
                            string nomeFileFMS = PrepareName(Workbook.GetUsrConfigElement("formatoNomeFileFMS").Value, codRup: entita["CodiceRup"].ToString()) + "*.xml";
                            string pathFileFMS = PreparePath(Workbook.GetUsrConfigElement("pathExportFileFMS"));

                            string[] files = Directory.GetFiles(pathFileFMS, nomeFileFMS, SearchOption.TopDirectoryOnly);

                            if (files.Length == 0)
                            {
                                if (!Workbook.DaConsole && System.Windows.Forms.MessageBox.Show("File FMS non presente nell'area di rete. Continuare con l'invio?", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                                {
                                    interrupt = true;
                                    break;
                                }
                            }
                            foreach (string file in files)
                            {
                                attachments.Add(file);
                                //18/01/2017 FIX: non più utile per cambio requisiti
                                //File.Copy(file, Path.Combine(pathExport, file.Split('\\').Last()));
                            }
                                
                        }
                        entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_ALLEGATO_RS' AND IdApplicazione = " + Workbook.IdApplicazione;
                        if (entitaProprieta.Count > 0)
                        {
                            string nomeFileFMS = PrepareName(Workbook.GetUsrConfigElement("formatoNomeFileRS_TERNA").Value) + ".xml";
                            string pathFileFMS = PreparePath(Workbook.GetUsrConfigElement("pathExportFileRS"));

                            string[] files = Directory.GetFiles(pathFileFMS, nomeFileFMS, SearchOption.TopDirectoryOnly);

                            if (files.Length == 0)
                            {
                                if (!Workbook.DaConsole && System.Windows.Forms.MessageBox.Show("File Riserva Secondaria non presente nell'area di rete. Continuare con l'invio?", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                                {
                                    interrupt = true;
                                    break;
                                }
                            }

                            foreach (string file in files)
                            {
                                attachments.Add(file);
                                //18/01/2017 FIX: non più utile per cambio requisiti
                                //File.Copy(file, Path.Combine(pathExport, file.Split('\\').Last()));
                            }
                        }
                    }

                    if (!interrupt)
                    {
                        var config = Workbook.GetUsrConfigElement("destMailTest");
                        string mailTo = config.Test;
                        string mailCC = "";

                        if (Workbook.Ambiente == Simboli.PROD)
                        {
                            entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_MAIL_TO' AND IdApplicazione = " + Workbook.IdApplicazione;
                            mailTo = entitaProprieta[0]["Valore"].ToString();
                            entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_MAIL_CC' AND IdApplicazione = " + Workbook.IdApplicazione;
                            mailCC = entitaProprieta[0]["Valore"].ToString();
                        }

                        entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_CODICE_MAIL' AND IdApplicazione = " + Workbook.IdApplicazione;
                        string codUP = entitaProprieta[0]["Valore"].ToString();

                        config = Workbook.GetUsrConfigElement("oggettoMail");
                        string oggetto = config.Value.Replace("%COD%", codUP).Replace("%DATA%", Workbook.DataAttiva.ToString("dd-MM-yyyy")).Replace("%MSD%", Workbook.Mercato) + (hasVariations ? " - CON VARIAZIONI" : "");
                        config = Workbook.GetUsrConfigElement("messaggioMail");
                        string messaggio = config.Value;

                        messaggio = Regex.Replace(messaggio, @"^[^\S\r\n]+", "", RegexOptions.Multiline);

                        //TODO check se manda sempre con lo stesso account...
                        Outlook.Account senderAccount = outlook.Session.Accounts[1];
                        foreach (Outlook.Account account in outlook.Session.Accounts)
                        {
                            if (account.DisplayName == "Bidding")
                                senderAccount = account;
                        }
                        mail.SendUsingAccount = senderAccount;
                        mail.Subject = oggetto;
                        mail.Body = messaggio;
                        foreach (string dest in mailTo.Split(';'))
                            if(dest.Trim() != "")
                                mail.Recipients.Add(dest.Trim());
                        mail.CC = mailCC;

                        //aggiungo allegato XLS
                        foreach (string attachment in attachments)
                            mail.Attachments.Add(attachment);

                        mail.Send();
                    }
                    File.Delete(excelExport);
                    if (pathXmlOrcoExportFull != null)
                    {
                        //proviamo a lasciare il file come da richiesta!
                        //File.Delete(pathXmlOrcoExportFull);
                    }

                    return !interrupt;
                }
            }
            catch(Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "InvioProgrammi - Esporta.InvioMail: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                foreach (string file in attachments)
                    File.Delete(file);

                return false;
            }

            return false;
        }

        // 29/11/2016 metode to create an XML for ORCO (from excel sheet named "Iren Idro")
        protected bool CreaOutputXML(object siglaEntita, string exportPath, string fileName, DateTime dataRif)
        {
            try
            {
                string nomeFoglio = "Iren Idro";
                DefinedNames definedNames = new DefinedNames(nomeFoglio);
                Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];
                int oreGiorno = Date.GetOreGiorno(dataRif);

                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                object codiceRUP = "UP_ORCO_1";
                object companyName = "IREN ENERGIA SPA";
                object companyID = "OEIESRD";
                
                DataView entitaAzioneInformazione = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
                entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaEntitaRif = 'UP_OCX' AND  IdApplicazione = " + Workbook.IdApplicazione;

                XNamespace ns = XNamespace.Get("urn:XML-PIPE");
                XNamespace xsi = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");
                XNamespace xsd = XNamespace.Get("http://www.w3.org/2001/XMLSchema");
                XNamespace schemaLocation = XNamespace.Get("urn:XML-PIPE PIPEDocument.xsd");

                string referenceNumber = codiceRUP.ToString().Replace("_", "") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                XElement PIPEDocument = new XElement(ns + "PIPEDocument",
                        new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                        new XAttribute(XNamespace.Xmlns + "xsd", xsd),
                        new XAttribute("ReferenceNumber", referenceNumber.Length > 30 ? referenceNumber.Substring(0, 30) : referenceNumber),
                        new XAttribute("CreationDate", DateTime.Now.ToString("yyyyMMddHHmmss")),
                        new XAttribute("Version", "1.0"),
                        new XAttribute(xsi + "schemaLocation", schemaLocation),
                        new XElement(ns + "TradingPartnerDirectory",
                            new XElement(ns + "Recipient",
                                new XElement(ns + "TradingPartner",
                                    new XAttribute("PartnerType", "Operator"),
                                    new XElement(ns + "CompanyName", companyName),
                                    new XElement(ns + "CompanyIdentifier", companyID)
                                )
                            )
                        )
                    );

                XElement fifteenMinuteSchedule = new XElement(ns + "FifteenMinuteSchedule",
                                new XAttribute("MarketParticipantNumber", "OEIESRD"),
                                new XAttribute("Type", "Original"),
                                new XElement(ns + "Market", Workbook.Mercato),
                                new XElement(ns + "Date", dataRif.ToString("yyyyMMdd")),
                                new XElement(ns + "UnitReferenceNumber", codiceRUP)
                            );

                if (entitaAzioneInformazione.Count > 0)
                {
                    DataRowView info = entitaAzioneInformazione[0];
                    string gradino = Regex.Match(info["SiglaInformazione"].ToString(), @"\d+").Value;
                    object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"];
                    Range rng = definedNames.Get(siglaEntitaRif, "PROGRAMMAQ" + gradino + "_" + Workbook.Mercato)
                        .Extend(rowOffset: 4, colOffset: oreGiorno);

                    for (int i = 0; i < oreGiorno; i++)
                    {
                        XElement hourDetail = new XElement(ns + "HourDetail",
                                new XElement(ns + "Hour", i + 1)
                            );

                        for (int quarter = 0; quarter < 4; quarter++)
                        {
                            /************************** 01/02/2017 **************************/
                            /***** Modifica valori espressi nel fiile xml  ****/
                            /************************** 01/02/2017 **************************/ 
 
                            /*
                            XElement quarterElement = new XElement(ns + "Quantity",
                                    new XAttribute("Minute", "0"),
                                    new XAttribute("QuarterInterval", quarter + 1),
                                    GetDecimal(ws, rng.Columns[i].Rows[quarter])
                                );
                             */
                            XElement quarterElement = new XElement(ns + "Quantity",
                                    new XAttribute("Minute", "0"),
                                    new XAttribute("QuarterInterval", quarter + 1),
                                    (GetDecimal(ws, rng.Columns[i].Rows[quarter])) / 4
                                );

                            hourDetail.Add(quarterElement);
                        }
                        fifteenMinuteSchedule.Add(hourDetail);
                    }
                    PIPEDocument.Add(new XElement(ns + "PIPTransaction", fifteenMinuteSchedule));
                }
                XDocument programmazioneImpianti = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
                        PIPEDocument
                    );
                programmazioneImpianti.Save(Path.Combine(exportPath, fileName));
                return true;
            }
            catch
            {
                return false;
            }
        }

        // 10/01/2017 metod to create a CSV file for ORCO (from excel sheet named "Iren Idro")
        protected bool CreaOutputCSV(object siglaEntita, string exportPath, string filename, DateTime dataRif)
        {
            try
            {
                string nomeFoglio = "Iren Idro";
                DefinedNames definedNames = new DefinedNames(nomeFoglio);
                Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];
                int oreGiorno = Date.GetOreGiorno(dataRif);

                string[] lines = new string[oreGiorno];

                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;
               
                object codiceRUP = "UP_OCX";
                object companyName = "IREN ENERGIA SPA";
                object companyID = "OEIESRD";
                
                DataView entitaAzioneInformazione = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
                entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaEntitaRif = 'UP_OCX' AND  IdApplicazione = " + Workbook.IdApplicazione;

                string referenceNumber = codiceRUP.ToString().Replace("_", "") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                if (entitaAzioneInformazione.Count > 0)
                {
                    DataRowView info = entitaAzioneInformazione[0];
                    string gradino = Regex.Match(info["SiglaInformazione"].ToString(), @"\d+").Value;
                    object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"];
                    Range rng = definedNames.Get(siglaEntitaRif, "PROGRAMMAQ" + gradino + "_" + Workbook.Mercato)
                        .Extend(rowOffset: 4, colOffset: oreGiorno);
   
                    for (int i = 0; i < oreGiorno; i++)
                    {
                        //Market
                        string str=Workbook.Mercato + ";";
                        //UnitReferenceNumber 
                        str += codiceRUP + ";";
                        //Date
                        str += dataRif.ToString("yyyyMMdd") + ";";
                        //Hour 
                        str += (i+1);
                        for (int quarter = 0; quarter < 4; quarter++)
                        {
                            //Quantity 1..4
                            str +=  ";" + GetDecimal(ws, rng.Columns[i].Rows[quarter]);
                        }
                        lines[i] = str;
                    }
                    File.WriteAllLines(Path.Combine(exportPath, filename), lines);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
