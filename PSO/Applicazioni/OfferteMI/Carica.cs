using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Iren.PSO.Applicazioni
{
    class Carica : Base.Carica
    {
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
                        DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno, parametro: Workbook.Mercato);
                    }
                    else
                    {
                        DataTable azioneInformazione = null;
                        if (mercati != null)
                            azioneInformazione = DataBase.Select(DataBase.SP.CARICA_AZIONE_INFORMAZIONE, "@SiglaEntita=" + siglaEntita + ";@SiglaAzione=" + siglaAzione + ";@Parametro=" + String.Join(",", mercati) + ";@Data=" + giorno.ToString("yyyyMMdd"));
                        else
                            azioneInformazione = DataBase.Select(DataBase.SP.CARICA_AZIONE_INFORMAZIONE, "@SiglaEntita=" + siglaEntita + ";@SiglaAzione=" + siglaAzione + ";@Parametro=" + parametro + ";@Data=" + giorno.ToString("yyyyMMdd"));

                        if (azioneInformazione != null)
                        {
                            if (azioneInformazione.Rows.Count == 0)
                            {
                                DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno, false, Workbook.Mercato);
                                return false;
                            }
                            else
                            {
                                ScriviInformazione(siglaEntita, azioneInformazione.DefaultView, definedNames);
                                DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno, parametro: Workbook.Mercato);
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
    }
}
