using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Iren.PSO.Base
{
    public class Date
    {
        #region Proprietà

        /// <summary>
        /// Scorciatoia per ottenere il suffisso della dataAttiva.
        /// </summary>
        public static string SuffissoDATA1
        {
            get { return GetSuffissoData(Workbook.DataAttiva); }
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Restituisce le ore di intervallo tra la data attiva e la data fine specificata.
        /// </summary>
        /// <param name="fine">Data fine.</param>
        /// <returns>Ore di intervallo.</returns>
        public static int GetOreIntervallo(DateTime fine)
        {
            return GetOreIntervallo(Workbook.DataAttiva, fine);
        }
        /// <summary>
        /// Restituisce le ore di intervallo tra una data inizio e fine specificate.
        /// </summary>
        /// <param name="inizio">Data inizio.</param>
        /// <param name="fine">Data fine.</param>
        /// <returns>Ore di intervallo.</returns>
        public static int GetOreIntervallo(DateTime inizio, DateTime fine)
        {
            return (int)(fine.AddDays(1).ToUniversalTime() - inizio.ToUniversalTime()).TotalHours;
        }
        /// <summary>
        /// Restituisce le ore che compongono il giorno passato per parametro.
        /// </summary>
        /// <param name="giorno">Giorno.</param>
        /// <returns>Numero di ore del giorno.</returns>
        public static int GetOreGiorno(DateTime giorno)
        {
            DateTime giornoSucc = giorno.AddDays(1);
            return (int)(giornoSucc.ToUniversalTime() - giorno.ToUniversalTime()).TotalHours;
        }
        /// <summary>
        /// Restituisce le ore che compongono il giorno passato per parametro.
        /// </summary>
        /// <param name="suffissoData">Suffisso del giorno.</param>
        /// <returns>Numero di ore del giorno.</returns>
        public static int GetOreGiorno(string suffissoData)
        {
            return GetOreGiorno(GetDataFromSuffisso(suffissoData));
        }
        /// <summary>
        /// Restituisce il suffisso del giorno rispetto alla data attiva.
        /// </summary>
        /// <param name="giorno">Giorno di cui trovare il suffisso.</param>
        /// <returns>Stringa del tipo DATAx con x = 1 se giorno è data attiva, x = 2 se giorno è data attiva + 1, e così via.</returns>
        public static string GetSuffissoData(DateTime giorno)
        {
            return GetSuffissoData(Workbook.DataAttiva, giorno);
        }
        /// <summary>
        /// Restituisce il suffisso del giorno rispetto alla data attiva.
        /// </summary>
        /// <param name="giorno">Giorno di cui trovare il suffisso.</param>
        /// <returns>Stringa del tipo DATAx con x = 1 se giorno è data attiva, x = 2 se giorno è data attiva + 1, e così via.</returns>
        public static string GetSuffissoData(string giorno)
        {
            return GetSuffissoData(Workbook.DataAttiva, giorno);
        }
        /// <summary>
        /// Restituisce il suffisso del giorno rispetto alla data inizio.
        /// </summary>
        /// <param name="inizio">Data di inizio.</param>
        /// <param name="giorno">Giorno di cui trovare il suffisso.</param>
        /// <returns>Stringa del tipo DATAx con x = 1 se giorno è data attiva, x = 2 se giorno è data attiva + 1, e così via.</returns>
        public static string GetSuffissoData(DateTime inizio, DateTime giorno)
        {
            if (inizio > giorno)
            {
                return "DATA0";
            }
            TimeSpan dayDiff = giorno - inizio;
            return "DATA" + (dayDiff.Days + 1);
        }
        /// <summary>
        /// Restituisce il suffisso del giorno rispetto alla data inizio.
        /// </summary>
        /// <param name="inizio">Data di inizio.</param>
        /// <param name="giorno">Giorno di cui trovare il suffisso.</param>
        /// <returns>Stringa del tipo DATAx con x = 1 se giorno è data attiva, x = 2 se giorno è data attiva + 1, e così via.</returns>
        public static string GetSuffissoData(DateTime inizio, object giorno)
        {
            DateTime day = DateTime.ParseExact(giorno.ToString().Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);
            return GetSuffissoData(inizio, day);
        }
        /// <summary>
        /// Restituisce il suffisso dell'ora in ingresso.
        /// </summary>
        /// <param name="ora">Numero rappresentante l'ora da 1 a 25.</param>
        /// <returns>Stringa del tipo Hx con x = ora.</returns>
        public static string GetSuffissoOra(int ora)
        {
            return "H" + ora;
        }
        /// <summary>
        /// Restituisce il suffisso dell'ora estraendolo dalla data ISO yyyyMMddHH
        /// </summary>
        /// <param name="dataOra">Stringa nella forma DATAx.Hy.</param>
        /// <returns>Stringa del tipo Hx.</returns>
        public static string GetSuffissoOra(object dataOra)
        {
            string dtO = dataOra.ToString();
            if (dtO.Length != 10)
                return "";

            return GetSuffissoOra(int.Parse(dtO.Substring(dtO.Length - 2, 2)));
        }
        /// <summary>
        /// Restituisce la data in formato ISO yyyyMMddHH a partire dal suffisso data e suffisso ora.
        /// </summary>
        /// <param name="data">Suffisso data.</param>
        /// <param name="ora">Suffisso ora.</param>
        /// <returns>Data in formato ISO yyyyMMddHH.</returns>
        public static string GetDataFromSuffisso(string data, string ora)
        {
            DateTime outDate = GetDataFromSuffisso(data);
            ora = ora == "" ? "0" : ora;
            int outOra = int.Parse(Regex.Match(ora, @"\d+").Value);

            return outDate.ToString("yyyyMMdd") + (outOra != 0 ? outOra.ToString("D2") : "");
        }
        /// <summary>
        /// Restituisce la data in formato ISO yyyyMMdd a partire dal suffisso data.
        /// </summary>
        /// <param name="data">Suffisso data.</param>
        /// <returns>Data in formato ISO yyyyMMdd.</returns>
        public static DateTime GetDataFromSuffisso(string data)
        {
            int giorno = int.Parse(Regex.Match(data.ToString(), @"\d+").Value);
            return Workbook.DataAttiva.AddDays(giorno - 1);
        }
        /// <summary>
        /// Restituisce l'ora a partire dalla stringa in formato ISO yyyyMMddHH
        /// </summary>
        /// <param name="dataOra"></param>
        /// <returns></returns>
        public static int GetOraFromDataOra(string dataOra)
        {
            string dtO = dataOra.ToString();
            if (dtO.Length != 10)
                return -1;

            return int.Parse(dtO.Substring(dtO.Length - 2, 2));
        }
        /// <summary>
        /// Restituisce l'ora a partire dal suffisso ora del tipo Hx.
        /// </summary>
        /// <param name="suffissoOra">Suffisso ora.</param>
        /// <returns>Intero rappresentante l'ora (1 - 25).</returns>
        public static int GetOraFromSuffissoOra(string suffissoOra)
        {
            string match = Regex.Match(suffissoOra, @"\d+").Value;
            return int.Parse(match);
        }

        #endregion
    }
}
