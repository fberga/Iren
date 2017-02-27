using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
/********* Aggiunte classi di serializzazione xml *************/
using System.Xml.Serialization;
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
            string mercato = mercati[0];
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
                        if (!CreaOfferteSuggeriteXML(siglaEntita, siglaAzione, pathStr, dataRif, mercato))
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
                // Create an instance of the XmlSerializer class;
                // specify the type of object to be deserialized.
                XmlSerializer serializer = new XmlSerializer(typeof(BMTransactionSUG));
                BMTransactionSUG bmt = new BMTransactionSUG();
                if (Workbook.DaConsole)
                {
                    bmt.ApplySendAutomaticSpecified = true;
                    bmt.ApplySendAutomatic = YESNO.YES;
                    bmt.OperatorCreator = Workbook.NomeUtente;
                }
                else
                {
                    bmt.ApplySendAutomaticSpecified = false;
                }

                //schemalocation non viene creato nel file xml e non so come impstarlo
                // XNamespace schemaLocation = XNamespace.Get("urn:XML-BIDMGM BM_SuggestedOfferMSD.xsd");
                //Data e ora creazione Offerta

                bmt.DataCreazione = (DateTime.Now.Date).ToString("yyyyMMdd");
                bmt.OraCreazione = (DateTime.Now.TimeOfDay.TotalMilliseconds).ToString();
                //bmt.OraCreazione = (DateTime.Now.TimeOfDay).ToString("HHmmss");
                //Reference Number
                string referenceNumber = codiceRUP.ToString().Replace("_", "") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                bmt.ReferenceNumber = referenceNumber;

                /* XNamespace ns = XNamespace.Get("urn:XML-BIDMGM");
                XNamespace xsi = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");
                XNamespace schemaLocation = XNamespace.Get("urn:XML-BIDMGM BM_SuggestedOfferMSD.xsd");
                */
                //Dictionary<string, SpecMercato> Mi = new Dictionary<string, SpecMercato>();
                List<Tuple<string, TimeSpan, TimeSpan, int, bool>> mercatoMI = new List<Tuple<string, TimeSpan, TimeSpan, int, bool>>();
                mercatoMI =  Simboli.MercatiMI;
                //Oggetto offerta suggerita
                bmt.Suggested = new Suggested();
                // Numero delle Entità selezionate (in questo metodo sempre una) oggetto del mercato.
                //TODO INSERIRE TUTTE LE ENTITA' SELEZIONATE
                int numEntità = 1;
                // verifico il mercato MI e le ore di quel mercato
                //controllo se ci sono dei vincoli di orario
                // Prima ora di mercato
                int oraInizio = mercatoMI.Where(x => x.Item1 == "MI" + mercato).FirstOrDefault().Item4 - 1;
                int oraFine = oreGiorno;

                // Intervallo ore oggetto di mercato (fine-inizio)
                int intervalloOreMercato = oraFine-oraInizio;
                // Array delle offete di mercato della unità selezionata 
                bmt.Suggested.Coordinate = new SuggestedCoordinate[numEntità];
                    
                // Offeta di mercato di una UP/UC
                bmt.Suggested.Coordinate[0] = new SuggestedCoordinate();
                // Data oggetto dell'offerta
                bmt.Suggested.Coordinate[0].FlowDate = dataRif.ToString("yyyyMMdd");
                // ID dell'unità
                bmt.Suggested.Coordinate[0].IDUnit = codiceRUP.ToString();
                // Mercato
                ElencoMercatiEnergia EnumMercato = (ElencoMercatiEnergia)Enum.Parse(typeof(ElencoMercatiEnergia), "MI" + mercato);
                bmt.Suggested.Coordinate[0].Mercato = EnumMercato;
                Range rngAV = new Range();
                Range rngVe = new Range();
                Range rngVp = new Range();
                Range rngBi = new Range();
                string prezzo = "";
                string energia = "";
                string codBil = "";
                int sgId = 0;
                for (int k = 1; k < 5; k++)
                {
                    /**/
                    rngAV = definedNames.Get(siglaEntita, "OFFERTA_MI" + mercato + "_G" + k + "TIPO", suffissoData).Extend(colOffset: oreGiorno);
                    rngVe = definedNames.Get(siglaEntita, "OFFERTA_MI" + mercato + "_G" + k + "E", suffissoData).Extend(colOffset: oreGiorno);
                    rngVp = definedNames.Get(siglaEntita, "OFFERTA_MI" + mercato + "_G" + k + "P", suffissoData).Extend(colOffset: oreGiorno);
                    rngBi = definedNames.Get(siglaEntita, "OFFERTA_MI" + mercato + "_G" + k + "CB", suffissoData).Extend(colOffset: oreGiorno);
                    energia = "0";
                    prezzo = "0";
                    if (!ws.Range[rngVe.ToString()].EntireRow.Hidden)
                    {
                        sgId++;
                        switch (sgId)
                        {
                            case 1:
                                // Array offerte ACQ/VEN per UP/UC Primo gradino
                                bmt.Suggested.Coordinate[0].SG1 = new SuggestedCoordinateSG1[intervalloOreMercato];
                                for (int j = oraInizio; j < oraFine; j++)
                                {
                                    // Dettaglio ACQ/VEN per ora j
                                    bmt.Suggested.Coordinate[0].SG1[j] = new SuggestedCoordinateSG1();
                                    // Azione ACQ/VEN
//                                        bmt.Suggested.Coordinate[0].SG1[j].AZIONE = TipoAzione.ACQ;
                                    bmt.Suggested.Coordinate[0].SG1[j].AZIONE = (TipoAzione)Enum.Parse(typeof(TipoAzione), (ws.Range[rngAV.Columns[j].ToString()].Value ?? "0").ToString());
//                                        bmt.Suggested.Coordinate[0].SG1[j].AZIONESpecified = true;
                                    //Quantità
                                    energia = (ws.Range[rngVe.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                                    bmt.Suggested.Coordinate[0].SG1[j].QUA = energia;
                                    //Prezzo
                                    prezzo = (ws.Range[rngVp.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                                    bmt.Suggested.Coordinate[0].SG1[j].PRE = prezzo;
                                    //Codice Bilanciamento
                                    codBil = (ws.Range[rngBi.Columns[j].ToString()].Value ?? "").ToString();
                                    bmt.Suggested.Coordinate[0].SG1[j].BILANC = codBil;                   
                                    // Ora oggetto del mercato.
                                    bmt.Suggested.Coordinate[0].SG1[j].Value = (j+1).ToString();
                                }
                                break;
                            case 2:
                                // Array offerte ACQ/VEN per UP/UC Secondo gradino
                                bmt.Suggested.Coordinate[0].SG2 = new SuggestedCoordinateSG2[intervalloOreMercato];
                                for (int j = oraInizio; j < oraFine; j++)
                                {
                                    // Dettaglio ACQ/VEN per ora j
                                    bmt.Suggested.Coordinate[0].SG2[j] = new SuggestedCoordinateSG2();
                                    // Azione ACQ/VEN
//                                        bmt.Suggested.Coordinate[0].SG2[j].AZIONE = TipoAzione.VEN;
                                    bmt.Suggested.Coordinate[0].SG1[j].AZIONE = (TipoAzione)Enum.Parse(typeof(TipoAzione), (ws.Range[rngAV.Columns[j].ToString()].Value ?? "0").ToString());
                                    //bmt.Suggested.Coordinate[0].SG2[i].AZIONESpecified = true;
                                    //Quantità
                                    energia = (ws.Range[rngVe.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                                    bmt.Suggested.Coordinate[0].SG2[j].QUA = energia;
                                    //Prezzo
                                    prezzo = (ws.Range[rngVp.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                                    bmt.Suggested.Coordinate[0].SG2[j].PRE = prezzo;
                                    //Codice Bilanciamento
                                    codBil = (ws.Range[rngBi.Columns[j].ToString()].Value ?? "").ToString();
                                    bmt.Suggested.Coordinate[0].SG2[j].BILANC = codBil;
                                    // Ora oggetto del mercato.
                                    bmt.Suggested.Coordinate[0].SG2[j].Value = (j + 1).ToString();
                                }
                                break;
                            case 3:
                                // Array offerte ACQ/VEN per UP/UC Terzo gradino
                                bmt.Suggested.Coordinate[0].SG3 = new SuggestedCoordinateSG3[intervalloOreMercato];
                                for (int j = oraInizio; j < oraFine; j++)
                                {
                                    // Dettaglio ACQ/VEN per ora j
                                    bmt.Suggested.Coordinate[0].SG3[j] = new SuggestedCoordinateSG3();
                                    // Azione ACQ/VEN
//                                        bmt.Suggested.Coordinate[0].SG3[j].AZIONE = TipoAzione.VEN;
                                    bmt.Suggested.Coordinate[0].SG1[j].AZIONE = (TipoAzione)Enum.Parse(typeof(TipoAzione), (ws.Range[rngAV.Columns[j].ToString()].Value ?? "0").ToString());
                                    // bmt.Suggested.Coordinate[0].SG3[i].AZIONESpecified = true;
                                    //Quantità
                                    energia = (ws.Range[rngVe.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                                    bmt.Suggested.Coordinate[0].SG3[j].QUA = energia;
                                    //Prezzo
                                    prezzo = (ws.Range[rngVp.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                                    bmt.Suggested.Coordinate[0].SG3[j].PRE = prezzo;
                                    //Codice Bilanciamento
                                    codBil = (ws.Range[rngBi.Columns[j].ToString()].Value ?? "").ToString();
                                    bmt.Suggested.Coordinate[0].SG3[j].BILANC = codBil;
                                    // Ora oggetto del mercato.
                                    bmt.Suggested.Coordinate[0].SG3[j].Value = (j + 1).ToString();
                                }
                                break;
                            case 4:
                                // Array offerte ACQ/VEN per UP/UC Quarto gradino
                                bmt.Suggested.Coordinate[0].SG4 = new SuggestedCoordinateSG4[intervalloOreMercato];
                                for (int j = oraInizio; j < oraFine; j++)
                                {
                                    // Dettaglio ACQ/VEN per ora j
                                    bmt.Suggested.Coordinate[0].SG4[j] = new SuggestedCoordinateSG4();
                                    // Azione ACQ/VEN
//                                        bmt.Suggested.Coordinate[0].SG4[j].AZIONE = TipoAzione.VEN;
                                    bmt.Suggested.Coordinate[0].SG1[j].AZIONE = (TipoAzione)Enum.Parse(typeof(TipoAzione), (ws.Range[rngAV.Columns[j].ToString()].Value ?? "0").ToString());
                                    //bmt.Suggested.Coordinate[0].SG4[j].AZIONESpecified = true;
                                    //Quantità
                                    energia = (ws.Range[rngVe.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                                    bmt.Suggested.Coordinate[0].SG4[j].QUA = energia;
                                    //Prezzo
                                    prezzo = (ws.Range[rngVp.Columns[j].ToString()].Value ?? "0").ToString().Replace(".", ",");
                                    bmt.Suggested.Coordinate[0].SG4[j].PRE = prezzo;
                                    //Codice Bilanciamento
                                    codBil = (ws.Range[rngBi.Columns[j].ToString()].Value ?? "").ToString();
                                    bmt.Suggested.Coordinate[0].SG4[j].BILANC = codBil;
                                    // Ora oggetto del mercato
                                    bmt.Suggested.Coordinate[0].SG4[j].Value = (j + 1).ToString();
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
                string filename = "Suggerite_MI_" + codiceRUP.ToString() + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + "_MI" + mercato + ".xml";
                TextWriter writer = new StreamWriter(Path.Combine(exportPath, filename));
                // Serialize the suggested offers.
                serializer.Serialize(writer, bmt);
                writer.Close();
             //   }
                    return true; 
            }
            catch
            {
                return false;
            }        
        }
    }
}
