using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Base
{
    /// <summary>
    /// Classe astratta che copre le funzionalità dell'azione di caricamento e generazione dei dati
    /// </summary>
    public abstract class ACarica
    {
        /// <summary>
        /// Launcher dell'azione di caricamento/generazione dei dati.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità di cui caricare/generare i dati.</param>
        /// <param name="siglaAzione">Sigla dell'azione per cui è richiesto il caricamento dei dati.</param>
        /// <param name="azionePadre">Sigla dell'azione padre (di solito CARICAx o GENERA).</param>
        /// <param name="giorno">Data di riferimento.</param>
        /// <param name="parametro">Parametro da specificare alla storedProcedure CARICA_AZIONE_INFORMAZIONE nel caso sia necessario.</param>
        /// <returns></returns>
        public abstract bool AzioneInformazione(object siglaEntita, object siglaAzione, object azionePadre, DateTime giorno, string[] mercati, object parametro = null);
    }

    /// <summary>
    /// Implementazione base della routine di caricamento/generazione delle informazioni in seguito ad un'azione dell'utente.
    /// </summary>
    public class Carica : ACarica
    {
        #region Metodi

        /// <summary>
        /// Launcher dell'azione di caricamento/generazione dei dati.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità di cui caricare/generare i dati.</param>
        /// <param name="siglaAzione">Sigla dell'azione per cui è richiesto il caricamento dei dati.</param>
        /// <param name="azionePadre">Sigla dell'azione padre (di solito CARICAx o GENERA).</param>
        /// <param name="giorno">Data di riferimento.</param>
        /// <param name="mercati">Mercati da considerare nell'azione.</param>
        /// <param name="parametro">Parametro da specificare alla storedProcedure CARICA_AZIONE_INFORMAZIONE nel caso sia necessario.</param>
        /// <returns>True se il caricamento va a buon fine.</returns>
        public override bool AzioneInformazione(object siglaEntita, object siglaAzione, object azionePadre, DateTime giorno, string[] mercati, object parametro = null)
        {
            DefinedNames definedNames = new DefinedNames(DefinedNames.GetSheetName(siglaEntita));
            try
            {
                AzzeraInformazione(siglaEntita, siglaAzione, definedNames, giorno, mercati, azionePadre.ToString().StartsWith("CARICA"));
                if (DataBase.OpenConnection())
                {
                    if (azionePadre.Equals("GENERA"))
                    {
                        if (mercati != null)
                        {
                            foreach (string mercato in mercati)
                            {
                                SpecMercato m = Simboli.MercatiMB["MB" + mercato];
                                ElaborazioneInformazione(siglaEntita, siglaAzione, definedNames, giorno, m.Inizio, Math.Min(Date.GetOreGiorno(giorno), m.Fine));
                            }
                        }
                        else
                            ElaborazioneInformazione(siglaEntita, siglaAzione, definedNames, giorno);
                        DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno);
                    }
                    else
                    {
                        DataTable azioneInformazione = null;
                        if(mercati != null)
                            azioneInformazione = DataBase.Select(DataBase.SP.CARICA_AZIONE_INFORMAZIONE, "@SiglaEntita=" + siglaEntita + ";@SiglaAzione=" + siglaAzione + ";@Parametro=" + String.Join(",",mercati) + ";@Data=" + giorno.ToString("yyyyMMdd"));
                        else
                            azioneInformazione = DataBase.Select(DataBase.SP.CARICA_AZIONE_INFORMAZIONE, "@SiglaEntita=" + siglaEntita + ";@SiglaAzione=" + siglaAzione + ";@Parametro=" + parametro + ";@Data=" + giorno.ToString("yyyyMMdd"));

                        if (azioneInformazione != null)
                        {
                            if (azioneInformazione.Rows.Count == 0)
                            {
                                DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno, false);
                                return false;
                            }
                            else
                            {
                                ScriviInformazione(siglaEntita, azioneInformazione.DefaultView, definedNames);
                                DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno);
                            }
                        }
                    }

                    Sheet s = new Sheet(Workbook.Sheets[definedNames.Sheet]);
                    s.AggiornaGrafici();
                    return true;
                }
                else
                {
                    if (azionePadre.Equals("GENERA")) 
                    {
                        ElaborazioneInformazione(siglaEntita, siglaAzione, definedNames, giorno);

                        Sheet s = new Sheet(Workbook.Sheets[definedNames.Sheet]);
                        s.AggiornaGrafici();

                        return true;
                    }

                    return false;
                }
            }
            catch (Exception e)
            {
                Workbook.InsertLog(PSO.Core.DataBase.TipologiaLOG.LogErrore, "CaricaAzioneInformazione [" + siglaEntita + ", " + siglaAzione + "]: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }
        /// <summary>
        /// Funzione che pulisce i campi di lavoro necessari al caricamento/generazione.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità di cui caricare/generare i dati.</param>
        /// <param name="siglaAzione">Sigla dell'azione per cui sono richieste la generazione o il caricamento dei dati.</param>
        /// <param name="definedNames">Oggetto che contiene l'indirizzamento delle celle per il foglio su cui si sta lavorando.</param>
        /// <param name="giorno">Data di riferimento.</param>
        protected virtual void AzzeraInformazione(object siglaEntita, object siglaAzione, DefinedNames definedNames, DateTime giorno, string[] mercati, bool isCarica)
        {
            Excel.Worksheet ws = Workbook.Sheets[definedNames.Sheet];

            string suffissoData = Date.GetSuffissoData(giorno);

            DataView azioneInformazione = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
            azioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            foreach (DataRowView info in azioneInformazione)
            {
                //cancella tutte le informazioni collegate all'azione che non contengono formule
                if (info["FormulaInCella"].Equals("0"))
                {
                    siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                    Range rng;
                    if (Base.Struct.tipoVisualizzazione == "R")
                    {
                        rng = definedNames.Get(siglaEntita, info["SiglaInformazione"], suffissoData).Extend(Struct.intervalloGiorni + 1, 25);
                    }
                    else
                    {
                        rng = definedNames.Get(siglaEntita, info["SiglaInformazione"], suffissoData).Extend(colOffset: Date.GetOreGiorno(giorno));
                    }

                    if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1"))
                    {
                        foreach (string mercato in mercati) 
                        {
                            Range mbRng = Simboli.GetMarketCompleteRange("MB" + mercato, giorno, rng);
                            if (mbRng != null)
                            {
                                ws.Range[mbRng.ToString()].Value = null;

                                if (!isCarica)
                                    Handler.StoreEdit(ws.Range[mbRng.ToString()], 0, true);
                                Style.RangeStyle(ws.Range[mbRng.ToString()], backColor: info["BackColor"], foreColor: info["ForeColor"]);
                                ws.Range[mbRng.ToString()].ClearComments();
                            }
                        }
                    }
                    else
                    {
                        ws.Range[rng.ToString()].Value = null;

                        if (!isCarica)
                            Handler.StoreEdit(ws.Range[rng.ToString()], 0, true);
                        Style.RangeStyle(ws.Range[rng.ToString()], backColor: info["BackColor"], foreColor: info["ForeColor"]);
                        ws.Range[rng.ToString()].ClearComments();
                    } 
                }
            }
        }
        /// <summary>
        /// Scrive le informazioni reperite dal database.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entita di cui sono state caricate le informazioni.</param>
        /// <param name="azioneInformazione">DataView contenente tutte le informazioni da inserire.</param>
        /// <param name="definedNames">Oggetto che contiene l'indirizzamento delle celle per il foglio su cui si sta lavorando.</param>
        protected virtual void ScriviInformazione(object siglaEntita, DataView azioneInformazione, DefinedNames definedNames)
        {
            Excel.Worksheet ws = Workbook.Sheets[definedNames.Sheet];

            foreach (DataRowView azione in azioneInformazione)
            {
                string suffissoData;
                string suffissoOra;
                if (azione["SiglaEntita"].Equals("UP_BUS") && azione["SiglaInformazione"].Equals("VOL_INVASO"))
                {
                    suffissoData = Date.GetSuffissoData(Workbook.DataAttiva.AddDays(-1));
                    suffissoOra = Date.GetSuffissoOra(24);
                }
                else
                {
                    suffissoData = Date.GetSuffissoData(Workbook.DataAttiva, azione["Data"]);
                    suffissoOra = Date.GetSuffissoOra(azione["Data"]);
                }

                ScriviCella(ws, definedNames, azione["SiglaEntita"], azione, suffissoData, suffissoOra, azione["Valore"], false, true);
            }
        }
        /// <summary>
        /// Funzione che elabora le informazioni correlate ad un azione per restituire il risultato richiesto.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità di cui generare i dati.</param>
        /// <param name="siglaAzione">Sigla dell'azione per cui è richiesta la generazione dei dati.</param>
        /// <param name="definedNames">Oggetto che contiene l'indirizzamento delle celle per il foglio su cui si sta lavorando.</param>
        /// <param name="giorno">Data di riferimento.</param>
        /// <param name="oraInizio">Vincolo sull'orario di inizio della generazione</param>
        /// <param name="oraFine">Vincolo sull'orario di fine della generazione</param>
        protected void ElaborazioneInformazione(object siglaEntita, object siglaAzione, DefinedNames definedNames, DateTime giorno, int oraInizio = -1, int oraFine = -1)
        {
            Excel.Worksheet ws = Workbook.Sheets[definedNames.Sheet];

            Dictionary<string, int> entitaRiferimento = new Dictionary<string, int>();
            List<int> oreDaCalcolare = new List<int>();

            string suffissoData = Date.GetSuffissoData(giorno);

            //controllo se ci sono dei vincoli di orario
            oraInizio = oraInizio < 0 ? 1 : oraInizio;
            oraFine = oraFine < 0 ? Date.GetOreGiorno(giorno) : oraFine;

            //cerco le entita che appartengono a quella in input
            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "Gerarchia = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            //salvo il numero di riferimento
            foreach (DataRowView entita in categoriaEntita)
                entitaRiferimento.Add(entita["SiglaEntita"].ToString(), (int)entita["Riferimento"]);

            if (entitaRiferimento.Count == 0)
                entitaRiferimento.Add(siglaEntita.ToString(), 1);

            DataView calcoloInformazione = Workbook.Repository[DataBase.TAB.CALCOLO_INFORMAZIONE].DefaultView;

            DataView entitaAzioneCalcolo = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE_CALCOLO].DefaultView;
            entitaAzioneCalcolo.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            foreach (DataRowView azioneCalcolo in entitaAzioneCalcolo)
            {
                calcoloInformazione.RowFilter = "SiglaCalcolo = '" + azioneCalcolo["SiglaCalcolo"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                calcoloInformazione.Sort = "Step";

                //azzero tutte le informazioni che vengono utilizzate nel calcolo tranne i CHECK
                foreach (DataRowView info in calcoloInformazione)
                {
                    if (!info["SiglaInformazione"].Equals("CHECKINFO"))
                    {
                        Range rng = definedNames.Get(info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"], info["SiglaInformazione"], suffissoData, Date.GetSuffissoOra(oraInizio)).Extend(colOffset: oraFine - oraInizio + 1);
                        ws.Range[rng.ToString()].Value = null;
                    }
                }

                for (int ora = oraInizio; ora <= oraFine; ora++)
                {
                    int i = 0;
                    while (i < calcoloInformazione.Count)
                    {
                        DataRowView calcolo = calcoloInformazione[i];
                        if (calcolo["OraInizio"] != DBNull.Value)
                            if (ora < int.Parse(calcolo["OraInizio"].ToString()) || ora > int.Parse(calcolo["OraFine"].ToString()))
                            {
                                i++;
                                continue;
                            }

                        if (calcolo["OraFine"] != DBNull.Value)
                            if (ora != Date.GetOreGiorno(giorno))
                                if (calcolo["FineCalcolo"].Equals("1"))
                                {
                                    i++;
                                    continue;
                                }
                                else
                                    break;

                        int step = 0;
                        object risultato = GetRisultatoCalcolo(siglaEntita, definedNames, giorno, ora, calcolo, entitaRiferimento, out step);

                        if (step == 0)
                        {
                            ScriviCella(ws, definedNames, siglaEntita, calcolo, suffissoData, Date.GetSuffissoOra(ora), risultato, true, false);
                        }

                        if (calcolo["FineCalcolo"].Equals("1") || step == -1)
                            break;

                        if (calcolo["GoStep"] != DBNull.Value)
                            step = (int)calcolo["GoStep"];

                        if (step != 0)
                            i = calcoloInformazione.Find(step);
                        else
                            i++;
                    }
                }
            }
        }
        /// <summary>
        /// Funzione che scrive le informazioni nella cella indicata dai parametri in input.
        /// </summary>
        /// <param name="ws">Foglio a cui appartengono le celle da scrivere.</param>
        /// <param name="definedNames">Oggetto che contiene l'indirizzamento delle celle per il foglio su cui si sta lavorando.</param>
        /// <param name="siglaEntita">Sigla dell'entità su cui scrivere i dati.</param>
        /// <param name="info">DataRow contenente le informazioni da scrivere nella cella</param>
        /// <param name="suffissoData">Suffisso della data di riferimento necessario per l'indirizzamento.</param>
        /// <param name="suffissoOra">Suffisso dell'ora di riferimento necessario per l'indirizzamento.</param>
        /// <param name="risultato">Risultato del calcolo da scrivere nella cella.</param>
        /// <param name="saveToDB">Flag che indica se l'informazione deve essere salvata o no sul DB in modo da attivare la routine di salvataggio della modifica.</param>
        protected virtual void ScriviCella(Excel.Worksheet ws, DefinedNames definedNames, object siglaEntita, DataRowView info, string suffissoData, string suffissoOra, object risultato, bool saveToDB, bool fromCarica) 
        {
            object siglaEntitaRif = siglaEntita;

            if(info.DataView.Table.Columns.Contains("SiglaEntitaRif") && info["SiglaEntitaRif"] != DBNull.Value)
                siglaEntitaRif = info["SiglaEntitaRif"];

            Range rng = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], suffissoData, suffissoOra);
            Excel.Range xlRng = ws.Range[rng.ToString()];

            xlRng.Value = risultato;

            if (info["BackColor"] != DBNull.Value)
                xlRng.Interior.ColorIndex = info["BackColor"];
            if (info["ForeColor"] != DBNull.Value)
                xlRng.Font.ColorIndex = info["ForeColor"];

            xlRng.ClearComments();

            if (info["Commento"] != DBNull.Value)
                xlRng.AddComment(info["Commento"]).Visible = false;

            if (saveToDB && !fromCarica)
                Handler.StoreEdit(xlRng, 0, true);
        }
        /// <summary>
        /// Funzione che esegue step by step il calcolo.
        /// </summary>
        /// <param name="siglaEntita">La sigla dell'entità da considerare.</param>
        /// <param name="definedNames">Oggetto che contiene l'indirizzamento delle celle per il foglio su cui si sta lavorando.</param>
        /// <param name="giorno">Data di riferimento.</param>
        /// <param name="ora">Ora di riferimento.</param>
        /// <param name="calcolo">Informazioni dello step corrente del calcolo.</param>
        /// <param name="entitaRiferimento">Struttura che contiene la lista delle entità di riferimento con annesso il codice del riferimento.</param>
        /// <param name="step">Lo step successivo a cui lo step corrente porta.</param>
        /// <returns>Il valore del calcolo effettuato.</returns>
        protected object GetRisultatoCalcolo(object siglaEntita, DefinedNames definedNames, DateTime giorno, int ora, DataRowView calcolo, Dictionary<string, int> entitaRiferimento, out int step)
        {
            Excel.Worksheet ws = Workbook.Sheets[definedNames.Sheet];

            string suffissoData = Date.GetSuffissoData(giorno);

            int ora1 = calcolo["OraInformazione1"] is DBNull ? ora : ora + (int)calcolo["OraInformazione1"];
            int ora2 = calcolo["OraInformazione2"] is DBNull ? ora : ora + (int)calcolo["OraInformazione2"];

            object siglaEntitaRif1 = calcolo["Riferimento1"] is DBNull ? (calcolo["SiglaEntita1"] is DBNull ? siglaEntita : calcolo["SiglaEntita1"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == (int)calcolo["Riferimento1"]).Key;
            object siglaEntitaRif2 = calcolo["Riferimento2"] is DBNull ? (calcolo["SiglaEntita2"] is DBNull ? siglaEntita : calcolo["SiglaEntita2"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == (int)calcolo["Riferimento2"]).Key;

            object valore1 = 0d;
            object valore2 = 0d;

            bool isEmptyVal1 = false;

            if (calcolo["SiglaInformazione1"] != DBNull.Value)
            {
                try
                {
                    Range cella1 = definedNames.Get(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1));
                    
                    isEmptyVal1 = ws.Range[cella1.ToString()].Value == null;

                    switch (calcolo["SiglaInformazione1"].ToString())
                    {
                        case "UNIT_COMM":
                            DataView entitaCommitment = Workbook.Repository[DataBase.TAB.ENTITA_COMMITMENT].DefaultView;
                            entitaCommitment.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' AND SiglaCommitment = '" + ws.Range[cella1.ToString()].Value + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                            valore1 = entitaCommitment.Count > 0 ? entitaCommitment[0]["IdEntitaCommitment"] : null;

                            break;
                        case "DISPONIBILITA":
                            if (ws.Range[cella1.ToString()].Value == "OFF")
                                valore1 = 0d;
                            else
                                valore1 = 1d;

                            break;
                        case "CHECKINFO":
                            if (ws.Range[cella1.ToString()].Value == "OK")
                                valore1 = 1d;
                            else
                                valore1 = 2d;
                            break;
                        default:
                            valore1 = ws.Range[cella1.ToString()].Value ?? 0d;
                            break;
                    }
                }
                catch
                {
                    valore1 = 0d;
                }
            }
            else if (calcolo["IdProprieta"] != DBNull.Value)
            {
                DataView entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' AND IdProprieta = " + calcolo["IdProprieta"] + " AND IdApplicazione = " + Workbook.IdApplicazione;

                if (entitaProprieta.Count > 0)
                    valore1 = entitaProprieta[0]["Valore"];
            }
            //else if (calcolo["IdParametroD"] != DBNull.Value)
            //{
            //    DataView entitaParametro = Workbook.Repository[DataBase.TAB.ENTITA_PARAMETRO_D].DefaultView;
            //    entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' AND IdParametro = " + calcolo["IdParametroD"] + " AND DataIV <= '" + Workbook.DataAttiva.ToString("yyyyMMdd") + "' AND DataFV > '" + Workbook.DataAttiva.ToString("yyyyMMdd") + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            //    if (entitaParametro.Count > 0)
            //        valore1 =entitaParametro[0]["Valore"];
            //}
            else if (calcolo["IdParametro"] != DBNull.Value)
            {
                DataView entitaParametro = Workbook.Repository[DataBase.TAB.ENTITA_PARAMETRO].DefaultView;
                entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' AND IdParametro = " + calcolo["IdParametro"] + " AND DataIV <= '" + Workbook.DataAttiva.ToString("yyyyMMdd") + ora1.ToString("00") + "' AND DataFV >= '" + Workbook.DataAttiva.ToString("yyyyMMdd") + ora1.ToString("00") + "' AND IdApplicazione = " + Workbook.IdApplicazione;

                if (entitaParametro.Count > 0)
                    valore1 = entitaParametro[0]["Valore"];
            }
            else if (calcolo["Valore"] != DBNull.Value)
            {
                valore1 = calcolo["Valore"];
            }

            if (calcolo["SiglaInformazione2"] != DBNull.Value)
            {
                try
                {
                    Range cella2 = definedNames.Get(siglaEntitaRif2, calcolo["SiglaInformazione2"], suffissoData, Date.GetSuffissoOra(ora2));

                    switch (calcolo["SiglaInformazione2"].ToString())
                    {
                        case "UNIT_COMM":
                            DataView entitaCommitment = Workbook.Repository[DataBase.TAB.ENTITA_COMMITMENT].DefaultView;
                            entitaCommitment.RowFilter = "SiglaEntita = '" + siglaEntitaRif2 + "' AND SiglaCommitment = '" + ws.Range[cella2.ToString()].Value + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                            valore2 = entitaCommitment.Count > 0 ? entitaCommitment[0] : null;

                            break;
                        case "DISPONIBILITA":
                            if (ws.Range[cella2.ToString()].Value == "OFF")
                                valore2 = 0d;
                            else
                                valore2 = 1d;

                            break;
                        case "CHECKINFO":
                            if (ws.Range[cella2.ToString()].Value == "OK")
                                valore2 = 1d;
                            else
                                valore2 = 2d;
                            break;
                        default:
                            valore2 = ws.Range[cella2.ToString()].Value ?? 0d;
                            break;
                    }
                }
                catch
                {
                    valore2 = 0d;
                }

            }

            double retVal = 0d;

            valore1 = valore1 ?? 0d;
            valore2 = valore2 ?? 0d;

            if (calcolo["Funzione"] is DBNull && calcolo["Operazione"] is DBNull && calcolo["Condizione"] is DBNull)
            {
                step = 0;
                if (Convert.ToDouble(valore1) == 0d)
                    return valore2;

                return valore1;
            }
            else if (calcolo["Funzione"] != DBNull.Value)
            {
                string func = calcolo["Funzione"].ToString().ToLowerInvariant();
                if (calcolo["SiglaInformazione2"] is DBNull)
                {
                    if (func.Contains("abs"))
                    {
                        retVal = Math.Abs(Convert.ToDouble(valore1));
                    }
                    else if (func.Contains("floor"))
                    {
                        retVal = Math.Floor(Convert.ToDouble(valore1));
                    }
                    else if (func.Contains("round"))
                    {
                        int decimals = int.Parse(func.Replace("round", ""));
                        retVal = Math.Round(Convert.ToDouble(valore1), decimals);
                    }
                    else if (func.Contains("power"))
                    {
                        int exp = int.Parse(Regex.Match(func, @"\d*").Value);
                        retVal = Math.Pow(Convert.ToDouble(valore1), exp);
                    }
                    else if (func.Contains("sum"))
                    {
                        foreach (var kvp in entitaRiferimento)
                            retVal += ws.Range[definedNames.Get(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0d;
                    }
                    else if (func.Contains("avg"))
                    {
                        foreach (var kvp in entitaRiferimento)
                            retVal += ws.Range[definedNames.Get(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0d;
                        retVal /= entitaRiferimento.Count;
                    }
                    else if (func.Contains("max_h"))
                    {
                        Range rng = definedNames.Get(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData).Extend(colOffset: Date.GetOreGiorno(giorno));
                        object[,] tmpVal = ws.Range[rng.ToString()].Value;
                        for (int i = 1; i <= tmpVal.GetLength(1); i++)
                            if (tmpVal[1, i] == null)
                                tmpVal[1, i] = 0d;

                        double[] values = tmpVal.Cast<double>().ToArray();
                        retVal = values.Max();
                    }
                    else if (func.Contains("min_h"))
                    {
                        Range rng = definedNames.Get(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData).Extend(colOffset: Date.GetOreGiorno(giorno));
                        object[,] tmpVal = ws.Range[rng.ToString()].Value;
                        for (int i = 1; i <= tmpVal.GetLength(1); i++)
                            if (tmpVal[1, i] == null)
                                tmpVal[1, i] = 0d;

                        double[] values = tmpVal.Cast<double>().ToArray();
                        retVal = values.Min();
                    }
                    else if (func.Contains("max"))
                    {
                        retVal = double.MinValue;
                        foreach (var kvp in entitaRiferimento)
                            retVal = Math.Max(ws.Range[definedNames.Get(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0, retVal);
                    }
                    else if (func.Contains("min"))
                    {
                        retVal = double.MaxValue;
                        foreach (var kvp in entitaRiferimento)
                            retVal = Math.Min(ws.Range[definedNames.Get(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0, retVal);
                    }
                    else if (func.Contains("isempty"))
                    {
                        retVal = isEmptyVal1 ? 1 : 0;
                    }
                }
                //caso in cui ci sia anche SiglaInformazione2
                else
                {
                    if (func.Contains("max"))
                    {
                        retVal = Math.Max(Convert.ToDouble(valore1), Convert.ToDouble(valore2));
                    }
                    else if (func.Contains("min"))
                    {
                        retVal = Math.Min(Convert.ToDouble(valore1), Convert.ToDouble(valore2));
                    }
                }
            }
            else if (calcolo["Operazione"] != DBNull.Value)
            {
                switch (calcolo["Operazione"].ToString())
                {
                    case "+":
                        retVal = Convert.ToDouble(valore1) + Convert.ToDouble(valore2);
                        break;
                    case "-":
                        retVal = Convert.ToDouble(valore1) - Convert.ToDouble(valore2);
                        break;
                    case "*":
                        retVal = Convert.ToDouble(valore1) * Convert.ToDouble(valore2);
                        break;
                    case "/":
                        retVal = Convert.ToDouble(valore1) / Convert.ToDouble(valore2);
                        break;
                }
            }
            else if (calcolo["Condizione"] != DBNull.Value)
            {
                bool res = false;
                switch (calcolo["Condizione"].ToString())
                {
                    case ">":
                        res = Convert.ToDouble(valore1) > Convert.ToDouble(valore2);
                        break;
                    case "<":
                        res = Convert.ToDouble(valore1) < Convert.ToDouble(valore2);
                        break;
                    case ">=":
                        res = Convert.ToDouble(valore1) >= Convert.ToDouble(valore2);
                        break;
                    case "<=":
                        res = Convert.ToDouble(valore1) <= Convert.ToDouble(valore2);
                        break;
                    case "=":
                        res = Convert.ToDouble(valore1) == Convert.ToDouble(valore2);
                        break;
                    case "<>":
                        res = Convert.ToDouble(valore1) != Convert.ToDouble(valore2);
                        break;
                }
                if (res)
                    step = (int)calcolo["StepCondizioneVera"];
                else
                    step = (int)calcolo["StepCondizioneFalsa"];

                return res;
            }

            step = 0;
            return retVal;
        }

        #endregion
    }
}
