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
                        if (!CreaOfferteSuggeriteXML(siglaEntita, siglaAzione, pathStr, dataRif, mercati))
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

        protected bool CreaOfferteSuggeriteXML(object siglaEntita, object siglaAzione, string exportPath, DateTime dataRif, string[] mercati)
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
                entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND idParametro = 903 AND CONVERT(DataIV, System.Int32) <= " + dataRif.ToString("yyyyMMdd") + " AND CONVERT(DataFV, System.Int32) >= " + dataRif.ToString("yyyyMMdd") + " AND IdApplicazione = " + Workbook.IdApplicazione;

                //decimal calcoloPPA = (decimal)entitaParametro[0]["Valore"];

                XNamespace ns = XNamespace.Get("urn:XML-BIDMGM");
                XNamespace xsi = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");
                XNamespace schemaLocation = XNamespace.Get("urn:XML-BIDMGM BM_SuggestedOfferMSD.xsd");

                string referenceNumber = codiceRUP.ToString().Replace("_", "") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                
                foreach (string mercato in mercati)
                {
                    XElement suggested = new XElement(ns + "Suggested");

                    // verifico il mercato MB e le ore di quel mercato
                    SpecMercato m = Simboli.MercatiMB["MB" + mercato];
                    //controllo se ci sono dei vincoli di orario
                    int oraInizio = m.Inizio - 1;
                    int oraFine = Math.Min(oreGiorno, m.Fine);


                    XElement coordinate = new XElement(ns + "Coordinate",
                        new XAttribute("Mercato", "MB" + mercato),
                        new XAttribute("IDUnit", codiceRUP),
                        new XAttribute("FlowDate", dataRif.ToString("yyyyMMdd"))
                    );

                    //cambioassetto
                    Range rng = new Range();
                    Range rng1 = new Range();
                    string prezzo = "";
                    string energia = "";
                    if (definedNames.TryGet(out rng, siglaEntita, "CAMBIO_ASSETTO_MB"))
                    {
                        //rng.StartColumn -= 1;
                        prezzo = (ws.Range[rng.ToString()].Value ?? "0").ToString().Replace(".", ",");
                        energia = "0";
                        XElement gradino = new XElement(ns + "CambioAssetto");
                        for (int j = oraInizio; j < oraFine; j++)
                            gradino.Add(new XElement(ns + "SG1", (j + 1),
                                    new XAttribute("PRE", prezzo),
                                    new XAttribute("QUA", energia),
                                    new XAttribute("AZIONE", "VEN")
                                )
                            );

                        coordinate.Add(gradino);
                    }

                    //spegnimento
                    rng = definedNames.Get(siglaEntita, "OFFERTA_MB_G0AE", suffissoData).Extend(colOffset: oreGiorno);
                    rng1 = definedNames.Get(siglaEntita, "OFFERTA_MB_G0AP", suffissoData).Extend(colOffset: oreGiorno);
                    energia = "0";
                    prezzo = "0";
                    if (!ws.Range[rng.ToString()].EntireRow.Hidden)
                    {
                        XElement gradino = new XElement(ns + "Spegnimento");
                        for (int j = oraInizio; j < oraFine; j++)
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
                    rng = definedNames.Get(siglaEntita, "OFFERTA_MB_G0VE", suffissoData).Extend(colOffset: oreGiorno);
                    rng1 = definedNames.Get(siglaEntita, "OFFERTA_MB_G0VP", suffissoData).Extend(colOffset: oreGiorno);
                    energia = "0";
                    prezzo = "0";
                    if (!ws.Range[rng.ToString()].EntireRow.Hidden)
                    {
                        XElement gradino = new XElement(ns + "Minimo");
                        for (int j = oraInizio; j < oraFine; j++)
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
                    rng = definedNames.Get(siglaEntita, "OFFERTA_MB_G4VE", suffissoData).Extend(colOffset: oreGiorno);
                    rng1 = definedNames.Get(siglaEntita, "OFFERTA_MB_G4VP", suffissoData).Extend(colOffset: oreGiorno);
                    energia = "0";
                    prezzo = "0";
                    if (!ws.Range[rng.ToString()].EntireRow.Hidden)
                    {
                        XElement gradino = new XElement(ns + "RisSecondaria");
                        for (int j = oraInizio; j < oraFine; j++)
                        {
                            energia = (ws.Range[rng.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                            prezzo = (ws.Range[rng1.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");

                            gradino.Add(new XElement(ns + "SG1", (j+1),
                                    new XAttribute("PRE", prezzo),
                                    new XAttribute("QUA", energia),
                                    new XAttribute("AZIONE", "VEN")
                                )
                            );
                        }

                        rng = definedNames.Get(siglaEntita, "OFFERTA_MB_G4AE", suffissoData).Extend(colOffset: oreGiorno);
                        rng1 = definedNames.Get(siglaEntita, "OFFERTA_MB_G4AP", suffissoData).Extend(colOffset: oreGiorno);
                        energia = "0";
                        prezzo = "0";

                        for (int j = oraInizio; j < oraFine; j++)
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
                        rng = definedNames.Get(siglaEntita, "OFFERTA_MB_G" + k + "VE", suffissoData).Extend(colOffset: oreGiorno);
                        rng1 = definedNames.Get(siglaEntita, "OFFERTA_MB_G" + k + "VP", suffissoData).Extend(colOffset: oreGiorno);
                        energia = "0";
                        prezzo = "0";
                        if (!ws.Range[rng.ToString()].EntireRow.Hidden)
                        {
                            aggiungi = true;
                            sgId++;
                            for (int j = oraInizio; j < oraFine; j++)
                            {
                                energia = (ws.Range[rng.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                                prezzo = (ws.Range[rng1.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");

                                XElement sg = new XElement(ns + ("SG" + sgId), (j + 1),
                                        new XAttribute("PRE", prezzo),
                                        new XAttribute("QUA", energia),
                                        new XAttribute("AZIONE", "VEN")
                                    );

                                /*if (calcoloPPA == 1 && k == 1 && j == 0)
                                    sg.Add(new XAttribute("RifStand", "MI1"));*/

                                altriServizi.Add(sg);
                            }

                            rng = definedNames.Get(siglaEntita, "OFFERTA_MB_G" + k + "AE", suffissoData).Extend(colOffset: oreGiorno);
                            rng1 = definedNames.Get(siglaEntita, "OFFERTA_MB_G" + k + "AP", suffissoData).Extend(colOffset: oreGiorno);
                            energia = "0";
                            prezzo = "0";
                            sgId++;
                            for (int j = oraInizio; j < oraFine; j++)
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
                    if (definedNames.TryGet(out rng, siglaEntita, "ACCENSIONE_MB"))
                    {
                        //AGGIUNTO FIL
                        rng = rng.Extend(colOffset: oreGiorno);
                        //prezzo = (ws.Range[rng.ToString()].Value ?? "0").ToString().Replace(".", ",");

                        XElement gradino = new XElement(ns + "Accensione");
                        for (int j = oraInizio; j < oraFine; j++)
                        {
                            //FIL
                            //energia = (ws.Range[rng.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");

                            object[,] prz = ws.Range[rng.Columns[j].ToString()].MergeArea.Value;

                            //TODO controllare se non esiste mergearea...va in errore!!

                            prezzo = (prz[1,1] ?? "0").ToString().Replace(".", ",");

                            gradino.Add(new XElement(ns + "SG1", (j + 1),
                                    new XAttribute("PRE", prezzo),
                                    new XAttribute("QUA", energia),
                                    new XAttribute("AZIONE", "VEN")
                                )
                            );
                        }

                        coordinate.Add(gradino);
                    }

                    suggested.Add(coordinate);



                    XDocument offerteSuggerite = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
                        new XElement(ns + "BMTransaction-SUGMSD",
                            new XAttribute("ReferenceNumber", referenceNumber.Length > 30 ? referenceNumber.Substring(0, 30) : referenceNumber),
                            new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                            new XAttribute(xsi + "schemaLocation", schemaLocation),
                            suggested
                        )
                    );
                    string filename = "Suggerite_MB_" + codiceRUP.ToString() + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + "_MB" + mercato + ".xml";
                    offerteSuggerite.Save(Path.Combine(exportPath, filename));
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
