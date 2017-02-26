using Iren.PSO.UserConfig;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace Iren.PSO.Base
{
    public abstract class AEsporta
    {
        #region Metodi

        /// <summary>
        /// Launcher per l'azione di esportazione, contiene il metodo standard di handling per eventuali errori. 
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità dell'export.</param>
        /// <param name="siglaAzione">Sigla dell'azione dell'export</param>
        /// <param name="desEntita">Descrizione dell'entità.</param>
        /// <param name="desAzione">Descrizione dell'azione.</param>
        /// <param name="dataRif">La data di riferimento per cui esportare i dati.</param>
        /// <returns>True se l'azione di esportazione è andata a buon fine, false altrimenti.</returns>
        public virtual bool RunExport(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif, string[] mercati)
        {
            try
            {
                if (EsportaAzioneInformazione(siglaEntita, siglaAzione, desEntita, desAzione, dataRif, mercati))
                {
                    if (DataBase.OpenConnection())
                        DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, dataRif);

                    DataBase.CloseConnection();

                    return true;
                }

                return false;
            }
            catch (Exception e)
            {
                if (DataBase.OpenConnection())
                    Workbook.InsertLog(PSO.Core.DataBase.TipologiaLOG.LogErrore, "RunExport [" + siglaEntita + ", " + siglaAzione + "]: " + e.Message);

                DataBase.CloseConnection();

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }
        /// <summary>
        /// Metodo virtuale di Esportazione.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità dell'export.</param>
        /// <param name="siglaAzione">Sigla dell'azione dell'export</param>
        /// <param name="desEntita">Descrizione dell'entità.</param>
        /// <param name="desAzione">Descrizione dell'azione.</param>
        /// <param name="dataRif">La data di riferimento per cui esportare i dati.</param>
        /// <returns></returns>
        protected abstract bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif, string[] mercati);
        /// <summary>
        /// Restituisce un'istanza di Outlook (quella aperta se ce n'è una, una nuova altrimenti).
        /// </summary>
        /// <returns>Istanza di Outlook.</returns>
        protected Outlook.Application GetOutlookInstance()
        {
            Outlook.Application application = null;
            try 
            {
                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            }
            catch
            {

                // If not, create a new instance of Outlook and log on to the default profile.
                application = new Outlook.Application();
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon("", "");
                nameSpace = null;
            }

            // Return the Outlook Application object.
            return application;
        }
        /// <summary>
        /// Metodo di esportazione su CSV. Scrive la tabella in ingresso in un file di testo situato al path indicato da nomeFile.
        /// </summary>
        /// <param name="nomeFile">Path del file.</param>
        /// <param name="dt">Tabella dei dati.</param>
        /// <returns>True se la scrittura ha avuto successo, false altrimenti.</returns>
        protected virtual bool ExportToCSV(string nomeFile, DataTable dt)
        {
            if (dt.Rows.Count > 0)
            {
                try
                {
                    using (StreamWriter outFile = new StreamWriter(nomeFile))
                    {
                        foreach (DataRow r in dt.Rows)
                        {
                            IEnumerable<string> fields = r.ItemArray.Select(field => field.ToString());
                            outFile.WriteLine(string.Join(";", fields));
                        }
                        outFile.Flush();
                    }
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            }

            return false;
        }

        #endregion

        #region Metodi Statici

        /// <summary>
        /// Prepara il path sostituendo le parti dinamiche con valori appropriati.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string PreparePath(UserConfigElement path)
        {
            //Controllo stato connessione
            string p;
            if(Workbook.Ambiente == Simboli.PROD)
                p = DataBase.OpenConnection() ? path.Value : path.Emergenza;
            else
                p = DataBase.OpenConnection() ? path.Test : path.Emergenza;
            DataBase.CloseConnection();

            return p;
        }

        public static string PrepareName(string name, string codRup = "")
        {
            Regex options = new Regex(@"\[\w+\]");
            name = options.Replace(name, match =>
            {
                string opt = match.Value.Replace("[", "").Replace("]", "");
                string o = "";
                switch (opt.ToLowerInvariant())
                {
                    case "msd":
                        o = Workbook.Mercato;
                        break;
                    case "codrup":
                        o = codRup;
                        break;
                    //aggiungere qui tutti i formati data da considerare nella forma
                    //case "formato data":
                    case "yyyymmdd":
                        o = Workbook.DataAttiva.ToString(opt);
                        break;
                }

                return o;
            });

            return name;
        }

        protected static decimal GetDecimal(Microsoft.Office.Interop.Excel.Worksheet ws, Range rngCorr, Range rng = null)
        {
            object val = null;
            
            if(rng == null)
                val = ws.Range[rngCorr.ToString()].Value ?? 0M;
            else
                val = ws.Range[rngCorr.ToString()].Value ?? ws.Range[rng.ToString()].Value ?? 0M;

            if (val.Equals(""))
                return 0M;

            return Convert.ToDecimal(val);
        }

        #endregion
    }

    public class Esporta : AEsporta
    {
        #region Metodi

        /// <summary>
        /// Metodo per eseguire un azione di esportazione. Da sovrascrivere in ogni applicativo che ha un'esportazione definita.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità dell'export.</param>
        /// <param name="siglaAzione">Sigla dell'azione dell'export</param>
        /// <param name="desEntita">Descrizione dell'entità.</param>
        /// <param name="desAzione">Descrizione dell'azione.</param>
        /// <param name="dataRif">La data di riferimento per cui esportare i dati.</param>
        /// <returns></returns>
        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif, string[] mercati)
        {
            return true;
        }

        #endregion
    }
}
