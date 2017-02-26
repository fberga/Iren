using Iren.PSO.Base;
using System;
using System.Data;
using System.IO;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Funzione di esportazione personalizzata.
    /// </summary>
    class Esporta : Base.Esporta
    {
        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif, string[] mercati)
        {
            DataView entitaAzione = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            if (entitaAzione.Count == 0)
                return false;

            switch (siglaAzione.ToString())
            {
                case "E_OFFERTA":

                    string pathStr = PreparePath(Workbook.GetUsrConfigElement("pathOfferteSuggerite"));
                    string emergenza = Workbook.GetUsrConfigElement("pathOfferteSuggerite").Emergenza;

                    if (Directory.Exists(pathStr))
                    {
                        if (!CreaOfferteSuggeriteXML_GME(siglaEntita, siglaAzione, emergenza, dataRif, "MSD1"))
                            return false;
                        if (!CreaOfferteSuggeriteXML(siglaEntita, siglaAzione, pathStr, dataRif, "MSD1"))
                            return false;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Il percorso '" + pathStr + "' non è raggiungibile.", Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                        return false;
                    }
                    
                    break;
            }
            return true;
        }

        protected bool CreaOfferteSuggeriteXML_GME(object siglaEntita, object siglaAzione, string exportPath, DateTime dataRif, string mercato)
        {
            try
            {
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                DefinedNames definedNames = new DefinedNames(nomeFoglio);
                Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];

                string suffissoData = Date.GetSuffissoData(dataRif);
                int oreGiorno = Date.GetOreGiorno(dataRif);

                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                object codiceRUP = categoriaEntita[0]["CodiceRUP"];
                //bool isTermo = categoriaEntita[0]["SiglaCategoria"].Equals("IREN_60T");

                DataView entitaParametro = Workbook.Repository[DataBase.TAB.ENTITA_PARAMETRO].DefaultView;
                entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND idParametro = 903 AND DataIV <= '" + dataRif.ToString("yyyyMMdd") + "01' AND DataFV >= '" + dataRif.ToString("yyyyMMdd") + "25' AND IdApplicazione = " + Workbook.IdApplicazione;

                decimal calcoloPPA = (decimal)entitaParametro[0]["Valore"];

                XNamespace ns = XNamespace.Get("urn:XML-PIPE");
                XNamespace xsi = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");
                XNamespace xsd = XNamespace.Get("http://www.w3.org/2001/XMLSchema");

                string referenceNumber = codiceRUP.ToString().Replace("_", "") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                DataView entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'COMPANY_NAME' AND IdApplicazione = " + Workbook.IdApplicazione;
                object companyName = entitaProprieta[0]["Valore"];

                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'COMPANY_IDENTIFIER' AND IdApplicazione = " + Workbook.IdApplicazione;
                object companyID = entitaProprieta[0]["Valore"];

                XElement PIPEDocument = new XElement(ns + "PIPEDocument",
                        new XAttribute("ReferenceNumber", referenceNumber.Length > 30 ? referenceNumber.Substring(0,30) : referenceNumber),
                        new XAttribute("CreationDate", DateTime.Now.ToString("yyyyMMddHHmmss")),
                        new XAttribute("Version", "1.0"),
                        new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                        new XAttribute(XNamespace.Xmlns + "xsd", xsd),
                        new XElement(ns + "TradingPartnerDirectory",
                            new XElement(ns + "Sender", 
                                new XElement(ns + "TradingPartner", 
                                    new XAttribute("PartnerType", "Market Participant"),
                                    new XElement(ns + "CompanyName", companyName),
                                    new XElement(ns + "CompanyIdentifier", companyID)
                                )
                            ),
                            new XElement(ns + "Recipient", 
                                new XElement(ns + "TradingPartner", 
                                    new XAttribute("PartnerType", "Operator"),
                                    new XElement(ns + "CompanyName", companyName),
                                    new XElement(ns + "CompanyIdentifier", companyID)
                                )
                            )
                        )
                    );

                string[] informazioni = { "OFFERTA_MSD_G0", "OFFERTA_MSD_G1", "OFFERTA_MSD_G2", "OFFERTA_MSD_G3", "OFFERTA_MSD_G4" };
                string[] gradini = { "AS", "GR1", "GR2", "GR3", "RS" };

                for (int i = 0; i < oreGiorno; i++)
                {
                    XElement bidSubmittal = new XElement(ns + "BidSubmittal",
                            new XAttribute("PredefinedOffer", "No"),                            
                            new XElement(ns + "Market", mercato),
                            new XElement(ns + "Date", dataRif.ToString("yyyyMMdd")),
                            new XElement(ns + "Hour", i + 1),
                            new XElement(ns + "UnitReferenceNumber", codiceRUP));

                    Range rng;
                    string presentedOffer;
                    string energia;
                    string prezzo;

                    for(int j = 0; j < informazioni .Length; j++)
                    {
                        //Vendita
                        rng = definedNames.Get(siglaEntita, informazioni[j]+"VE", suffissoData, Date.GetSuffissoOra(i + 1));
                        presentedOffer = "No";
                        energia = "0";
                        prezzo = "0";
                        if(!ws.Range[rng.ToString()].EntireRow.Hidden) 
                        {
                            presentedOffer = "Yes";
                            energia = (ws.Range[rng.ToString()].Value ?? "0").ToString().Replace(".", ",");

                            rng = definedNames.Get(siglaEntita, informazioni[j] + "VP", suffissoData, Date.GetSuffissoOra(i + 1));
                            if (ws.Range[rng.ToString()].Value != null)
                                prezzo = ws.Range[rng.ToString()].Value.ToString().Replace(".", ",");
                        }
                        
                        bidSubmittal.Add(new XElement(ns + "Offer",
                                new XAttribute("PresentedOffer", presentedOffer),
                                new XAttribute("Purpose", "Sell"),
                                new XAttribute("Scope", gradini[j]),
                                new XElement(ns + "BidQuantity", energia,
                                    new XAttribute("UnitOfMeasure", "MWh")),
                                new XElement(ns + "EnergyPrice", prezzo),
                                new XElement(ns + "SourceOffer", "SPOT"))
                            );
                        
                        //Acquisto
                        rng = definedNames.Get(siglaEntita, informazioni[j]+"AE", suffissoData, Date.GetSuffissoOra(i + 1));
                        presentedOffer = "No";
                        energia = "0";
                        prezzo = "0";
                        if(!ws.Range[rng.ToString()].EntireRow.Hidden) 
                        {
                            presentedOffer = "Yes";
                            energia = (ws.Range[rng.ToString()].Value ?? "0").ToString().Replace(".", ",");

                            rng = definedNames.Get(siglaEntita, informazioni[j] + "AP", suffissoData, Date.GetSuffissoOra(i + 1));

                            if (ws.Range[rng.ToString()].Value != null)
                                prezzo = ws.Range[rng.ToString()].Value.ToString().Replace(".", ",");
                        }

                        bidSubmittal.Add(new XElement(ns + "Offer",
                                new XAttribute("PresentedOffer", presentedOffer),
                                new XAttribute("Purpose", "Buy"),
                                new XAttribute("Scope", gradini[j]),
                                new XElement(ns + "BidQuantity", energia,
                                    new XAttribute("UnitOfMeasure", "MWh")),
                                new XElement(ns + "EnergyPrice", prezzo),
                                new XElement(ns + "SourceOffer", "SPOT"))
                            );
                    }

                    //Accensione - Vendita
                    presentedOffer = "Yes";
                    prezzo = "0";
                    energia = "0";
                    if (definedNames.TryGet(out rng, siglaEntita, "ACCENSIONE_MSD"))
                    {
                        //aggiusto la colonna che mi ritorna DATA1.H1
                        //rng.StartColumn -= 1;
                        if(ws.Range[rng.ToString()].Value != null)
                            prezzo = ws.Range[rng.ToString()].Value.ToString().Replace(".", ",");
                    }

                    bidSubmittal.Add(new XElement(ns + "Offer",
                            new XAttribute("PresentedOffer", presentedOffer),
                            new XAttribute("Purpose", "Sell"),
                            new XAttribute("Scope", "AC"),
                            new XElement(ns + "BidQuantity", energia,
                                new XAttribute("UnitOfMeasure", "MWh")),
                            new XElement(ns + "EnergyPrice", prezzo),
                            new XElement(ns + "SourceOffer", "SPOT"))
                        );

                    //Cambio Assetto - Vendita
                    presentedOffer = "Yes";
                    prezzo = "0";
                    if (definedNames.TryGet(out rng, siglaEntita, "CAMBIO_ASSETTO_MSD"))
                    {
                        //aggiusto la colonna che mi ritorna DATA1.H1
                        //rng.StartColumn -= 1;
                        if (ws.Range[rng.ToString()].Value != null)
                            prezzo = ws.Range[rng.ToString()].Value.ToString().Replace(".", ",");
                    }

                    bidSubmittal.Add(new XElement(ns + "Offer",
                            new XAttribute("PresentedOffer", presentedOffer),
                            new XAttribute("Purpose", "Sell"),
                            new XAttribute("Scope", "CA"),
                            new XElement(ns + "BidQuantity", energia,
                                new XAttribute("UnitOfMeasure", "MWh")),
                            new XElement(ns + "EnergyPrice", prezzo),
                            new XElement(ns + "SourceOffer", "SPOT"))
                        );

                    if(calcoloPPA == 1)
                        bidSubmittal.Add(new XAttribute("RifStand", "MI1"));

                    PIPEDocument.Add(new XElement(ns + "PIPTransaction", bidSubmittal));
                }

                XDocument offerteSuggerite = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
                        PIPEDocument
                    );

                string filename = "Suggerite_MSD_" + codiceRUP.ToString() + "_GME.xml";
                offerteSuggerite.Save(Path.Combine(exportPath, filename));

                return true;
            }
            catch
            {
                return false;
            }
        }
        protected bool CreaOfferteSuggeriteXML(object siglaEntita, object siglaAzione, string exportPath, DateTime dataRif, string mercato)
        {
            try
            {
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                DefinedNames definedNames = new DefinedNames(nomeFoglio);
                Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];

                string suffissoData = Date.GetSuffissoData(dataRif);
                int oreGiorno = Date.GetOreGiorno(dataRif);

                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                object codiceRUP = categoriaEntita[0]["CodiceRUP"];
                //bool isTermo = categoriaEntita[0]["SiglaCategoria"].Equals("IREN_60T");

                DataView entitaParametro = Workbook.Repository[DataBase.TAB.ENTITA_PARAMETRO].DefaultView;
                entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND idParametro = 903 AND DataIV <= '" + dataRif.ToString("yyyyMMdd") + "01' AND DataFV >= '" + dataRif.ToString("yyyyMMdd") + "25' AND IdApplicazione = " + Workbook.IdApplicazione;

                decimal calcoloPPA = (decimal)entitaParametro[0]["Valore"];

                XNamespace ns = XNamespace.Get("urn:XML-BIDMGM");
                XNamespace xsi = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");
                XNamespace schemaLocation = XNamespace.Get("urn:XML-BIDMGM BM_SuggestedOfferMSD.xsd");

                string referenceNumber = codiceRUP.ToString().Replace("_", "") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                XElement BMTransaction = new XElement(ns + "BMTransaction-SUGMSD",
                        new XAttribute("ReferenceNumber", referenceNumber.Length > 30 ? referenceNumber.Substring(0, 30) : referenceNumber),
                        new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                        new XAttribute(xsi + "schemaLocation", schemaLocation),
                        new XElement(ns + "Suggested")
                    );

                XElement coordinate = new XElement(ns + "Coordinate",
                    new XAttribute("Mercato", "MSD"),
                    new XAttribute("IDUnit", codiceRUP),
                    new XAttribute("FlowDate", dataRif.ToString("yyyyMMdd"))
                );

                //cambioassetto
                Range rng = new Range();
                Range rng1 = new Range();
                string prezzo = "";
                string energia = "";
                if (definedNames.TryGet(out rng, siglaEntita, "CAMBIO_ASSETTO_MSD"))
                {
                    //rng.StartColumn -= 1;
                    prezzo = (ws.Range[rng.ToString()].Value ?? "0").ToString().Replace(".", ",");
                    energia = "0";
                    XElement gradino = new XElement(ns + "CambioAssetto");
                    for (int j = 0; j < oreGiorno; j++)
                        gradino.Add(new XElement(ns + "SG1", (j + 1),
                                new XAttribute("PRE", prezzo),
                                new XAttribute("QUA", energia),
                                new XAttribute("AZIONE", "VEN")
                            )
                        );

                    coordinate.Add(gradino);
                }

                //spegnimento
                rng = definedNames.Get(siglaEntita, "OFFERTA_MSD_G0AE", suffissoData).Extend(colOffset: oreGiorno);
                rng1 = definedNames.Get(siglaEntita, "OFFERTA_MSD_G0AP", suffissoData).Extend(colOffset: oreGiorno);
                energia = "0";
                prezzo = "0";
                if (!ws.Range[rng.ToString()].EntireRow.Hidden)
                {
                    XElement gradino = new XElement(ns + "Spegnimento");
                    for (int j = 0; j < oreGiorno; j++)
                    {
                        energia = (ws.Range[rng.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                        prezzo = (ws.Range[rng1.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");

                        gradino.Add(new XElement(ns + "SG1", (j + 1),
                                new XAttribute("PRE", prezzo),
                                new XAttribute("QUA", energia),
                                new XAttribute("AZIONE", "ACQ")
                            )
                        );
                    }
                    coordinate.Add(gradino);
                }

                //minimo
                rng = definedNames.Get(siglaEntita, "OFFERTA_MSD_G0VE", suffissoData).Extend(colOffset: oreGiorno);
                rng1 = definedNames.Get(siglaEntita, "OFFERTA_MSD_G0VP", suffissoData).Extend(colOffset: oreGiorno);
                energia = "0";
                prezzo = "0";
                if (!ws.Range[rng.ToString()].EntireRow.Hidden)
                {
                    XElement gradino = new XElement(ns + "Minimo");
                    for (int j = 0; j < oreGiorno; j++)
                    {
                        energia = (ws.Range[rng.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                        prezzo = (ws.Range[rng1.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");

                        gradino.Add(new XElement(ns + "SG1", (j + 1),
                                new XAttribute("PRE", prezzo),
                                new XAttribute("QUA", energia),
                                new XAttribute("AZIONE", "VEN")
                            )
                        );
                    }
                    coordinate.Add(gradino);
                }

                //riserva secondaria
                rng = definedNames.Get(siglaEntita, "OFFERTA_MSD_G4VE", suffissoData).Extend(colOffset: oreGiorno);
                rng1 = definedNames.Get(siglaEntita, "OFFERTA_MSD_G4VP", suffissoData).Extend(colOffset: oreGiorno);
                energia = "0";
                prezzo = "0";
                if (!ws.Range[rng.ToString()].EntireRow.Hidden)
                {
                    XElement gradino = new XElement(ns + "RisSecondaria");
                    for (int j = 0; j < oreGiorno; j++)
                    {
                        energia = (ws.Range[rng.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                        prezzo = (ws.Range[rng1.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");

                        gradino.Add(new XElement(ns + "SG1", (j + 1),
                                new XAttribute("PRE", prezzo),
                                new XAttribute("QUA", energia),
                                new XAttribute("AZIONE", "VEN")
                            )
                        );
                    }

                    rng = definedNames.Get(siglaEntita, "OFFERTA_MSD_G4AE", suffissoData).Extend(colOffset: oreGiorno);
                    rng1 = definedNames.Get(siglaEntita, "OFFERTA_MSD_G4AP", suffissoData).Extend(colOffset: oreGiorno);
                    energia = "0";
                    prezzo = "0";
                    
                    for (int j = 0; j < oreGiorno; j++)
                    {
                        energia = (ws.Range[rng.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                        prezzo = (ws.Range[rng1.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");

                        gradino.Add(new XElement(ns + "SG2", (j + 1),
                                new XAttribute("PRE", prezzo),
                                new XAttribute("QUA", energia),
                                new XAttribute("AZIONE", "ACQ")
                            )
                        );
                    }

                    coordinate.Add(gradino);
                }

                //altri servizi
                XElement altriServizi = new XElement(ns + "AltriServizi");

                bool aggiungi = false;
                int sgId = 0;
                for (int k = 1; k < 4; k++)
                {
                    rng = definedNames.Get(siglaEntita, "OFFERTA_MSD_G" + k + "VE", suffissoData).Extend(colOffset: oreGiorno);
                    rng1 = definedNames.Get(siglaEntita, "OFFERTA_MSD_G" + k + "VP", suffissoData).Extend(colOffset: oreGiorno);
                    energia = "0";
                    prezzo = "0";
                    if (!ws.Range[rng.ToString()].EntireRow.Hidden)
                    {
                        aggiungi = true;
                        sgId++;
                        for (int j = 0; j < oreGiorno; j++)
                        {
                            energia = (ws.Range[rng.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                            prezzo = (ws.Range[rng1.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");

                            XElement sg = new XElement(ns + ("SG" + sgId), (j + 1),
                                    new XAttribute("PRE", prezzo),
                                    new XAttribute("QUA", energia),
                                    new XAttribute("AZIONE", "VEN")
                                );

                            if (calcoloPPA == 1 && k == 1 && j == 0)
                                sg.Add(new XAttribute("RifStand", "MI1"));

                            altriServizi.Add(sg);
                        }

                        rng = definedNames.Get(siglaEntita, "OFFERTA_MSD_G" + k + "AE", suffissoData).Extend(colOffset: oreGiorno);
                        rng1 = definedNames.Get(siglaEntita, "OFFERTA_MSD_G" + k + "AP", suffissoData).Extend(colOffset: oreGiorno);
                        energia = "0";
                        prezzo = "0";
                        sgId++;
                        for (int j = 0; j < oreGiorno; j++)
                        {
                            energia = (ws.Range[rng.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                            prezzo = (ws.Range[rng1.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");

                            altriServizi.Add(new XElement(ns + ("SG" + sgId), (j + 1),
                                    new XAttribute("PRE", prezzo),
                                    new XAttribute("QUA", energia),
                                    new XAttribute("AZIONE", "ACQ")
                                )
                            );
                        }
                    }
                }
                if(aggiungi)
                    coordinate.Add(altriServizi);

                //accensione
                rng = new Range();
                energia = "0";
                prezzo = "0";
                if (definedNames.TryGet(out rng, siglaEntita, "ACCENSIONE_MSD"))
                {
                    //rng.StartColumn -= 1;
                    prezzo = (ws.Range[rng.ToString()].Value ?? "0").ToString().Replace(".", ",");

                    XElement gradino = new XElement(ns + "Accensione");
                    for (int j = 0; j < oreGiorno; j++)
                    {
                        gradino.Add(new XElement(ns + "SG1", (j + 1),
                                new XAttribute("PRE", prezzo),
                                new XAttribute("QUA", energia),
                                new XAttribute("AZIONE", "VEN")
                            )
                        );
                    }

                    coordinate.Add(gradino);
                }

                XDocument offerteSuggerite = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
                        new XElement(ns + "BMTransaction-SUGMSD",
                            new XAttribute("ReferenceNumber", referenceNumber.Length > 30 ? referenceNumber.Substring(0, 30) : referenceNumber),
                            new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                            new XAttribute(xsi + "schemaLocation", schemaLocation),
                            new XElement(ns + "Suggested", coordinate)
                        )
                    );

                string filename = "Suggerite_MSD_" + codiceRUP.ToString() + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xml";
                offerteSuggerite.Save(Path.Combine(exportPath, filename));

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
