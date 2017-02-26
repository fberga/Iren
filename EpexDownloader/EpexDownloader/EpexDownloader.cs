using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using HtmlAgilityPack;
using System.Globalization;
using System.Data;
using System.IO;
using System.Configuration;

namespace Iren.EpexDownloader
{
    class EpexDownloader
    {
        #region Variabili

        private WebClient _webClient = new WebClient();
        private HtmlAgilityPack.HtmlDocument _htmlDoc = new HtmlAgilityPack.HtmlDocument();
        private string _baseURL = "http://www.epexspot.com/en/market-data/dayaheadauction/auction-table/";
        private string _basePath = @"D:\Users\e-bergamin\Desktop";
        private DateTime _dataInizio;
        private DateTime _dataFine;

        #endregion

        static void Main(string[] args)
        {
            EpexDownloader epexDwnloader = new EpexDownloader();

            for (; epexDwnloader._dataInizio <= epexDwnloader._dataFine; epexDwnloader._dataInizio = epexDwnloader._dataInizio.AddDays(1))
            {
                Console.Write("Data: " + epexDwnloader._dataInizio.ToString("dd/MM/yyyy") + "...");
                epexDwnloader.Run(epexDwnloader._dataInizio);
                Console.WriteLine(" Done!");
            }
            Console.WriteLine("Done");
        }

        #region Costruttori

        public EpexDownloader()
        {
            _basePath = ConfigurationManager.AppSettings["basePath"] ?? _basePath;
            _baseURL = ConfigurationManager.AppSettings["baseURL"] ?? _baseURL;

            if (!DateTime.TryParseExact(ConfigurationManager.AppSettings["endDate"], "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _dataFine))
            {
                _dataFine = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day).AddDays(1);
                
            }
            if (!DateTime.TryParseExact(ConfigurationManager.AppSettings["startDate"], "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _dataInizio))
            {
                _dataInizio = _dataFine;
            }
        }

        #endregion

        #region Metodi

        public void Run(DateTime day)
        {
            bool is25hours = (day.Month == 10 && isLastSunday(day));
            bool is23hours = !is25hours && (day.Month == 3 && isLastSunday(day));

            string URL = _baseURL + day.ToString("yyyy-MM-dd") + "/FR";
            try
            {
                _htmlDoc.LoadHtml(_webClient.DownloadString(URL));

                //ottengo l'array delle date visualizzate
                HtmlNode dateRow = _htmlDoc.DocumentNode.SelectSingleNode("//div[@id='tab_fr']//table[@class='list hours responsive']//tr");
                List<DateTime> days = new List<DateTime>();
                foreach (HtmlNode col in dateRow.SelectNodes("th"))
                {
                    DateTime d = new DateTime();
                    if (DateTime.TryParseExact(col.InnerText + " " + day.Year, "ddd, dd/MM yyyy", new CultureInfo("en-US"), DateTimeStyles.None, out d))
                        days.Add(d);
                }

                KeyValuePair<string, int>[] tabIDs = new KeyValuePair<string, int>[] 
                { 
                    new KeyValuePair<string, int>("tab_fr", 987), 
                    new KeyValuePair<string, int>("tab_de", 924), 
                    new KeyValuePair<string, int>("tab_ch", 988)};

                foreach (KeyValuePair<string, int> tabID in tabIDs)
                {
                    HtmlNodeCollection tab = _htmlDoc.DocumentNode.SelectNodes("//div[@id='" + tabID.Key + "']//table[@class='list hours responsive']//tr[@class='no-border']");

                    //la mia data ha 24 ore ma la tabella contiene anche la riga della 25-esima
                    if (!is25hours && tab.Count() == 25)
                        tab.RemoveAt(3);

                    DataTable dt = initTable();

                    int i = 0;
                    int index = days.IndexOf(day);
                    foreach (HtmlNode row in tab)
                    {
                        //seleziono il valore che mi interessa dalla tabella sapendo che index è 0-based e che le prime 2 colonne sono di intestazione
                        HtmlNode mgpVal = row.SelectSingleNode("td[" + (3 + index) + "]");
                        DataRow newRow = dt.NewRow();

                        newRow["Zona"] = tabID.Value;
                        newRow["Data"] = day.ToString("yyyyMMdd") + (++i < 10 ? "0" : "") + i;
                        newRow["Mgp"] = 0;
                        decimal tmp;
                        if (Decimal.TryParse(mgpVal.InnerText.Replace('.', ','), out tmp))
                            newRow["MGP"] = tmp;

                        dt.Rows.Add(newRow);
                    }

                    if (dt.Rows.Count > 0)
                    {
                        //scrivo la tabella all'interno del caricatore
                        string path = Path.Combine(_basePath, day.ToString("yyyyMMdd") + "_" + tabID.Value + ".xml");
                        dt.WriteXml(path);
                    }
                }
            }
            catch(Exception)
            {

            }
        }

        private DataTable initTable()
        {
            DataTable dt = new DataTable("Epex")
            {
                Columns =
                {
                    {"Zona", typeof(int)},
                    {"Data", typeof(string)},
                    {"Mgp", typeof(Decimal)}
                }
            };

            return dt;
        }

        private DateTime GetLastWeekdayOfMonth(DateTime date, DayOfWeek day)
        {
            DateTime lastDayOfMonth = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
            int wantedDay = (int)day;
            int lastDay = (int)lastDayOfMonth.DayOfWeek;
            return lastDayOfMonth.AddDays(
                lastDay >= wantedDay ? wantedDay - lastDay : wantedDay - lastDay - 7);
        }

        private Boolean isLastSunday(DateTime date)
        {
            return date == GetLastWeekdayOfMonth(date, DayOfWeek.Sunday);
        }

        #endregion


    }
}
