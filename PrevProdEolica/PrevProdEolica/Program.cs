using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace PrevProdEolica
{
    class Program
    {
        static void Main(string[] args)
        {
            Downloader dwnER = new Downloader();

            Console.WriteLine("Inizio download per la data -> {0} ...", dwnER.Data.ToShortDateString());
            dwnER.Run();
            Console.WriteLine("... done!");
            Console.Read();
        }
    }

    class Downloader
    {
        #region Variabili

        private WebClient _webClient = new WebClient();
        private HtmlDocument _htmlDoc = new HtmlDocument();
        private string _baseURL = "http://www.terna.it";
        private string _dwnldURL = "/default/Home/SISTEMA_ELETTRICO/transparency_report/Generation/Forecast_generation_wind.aspx";
        private string _basePath = @"D:\Users\e-bergamin\Desktop";
        private DateTime _data;

        #endregion

        #region Proprietà

        public DateTime Data { get { return _data; } }

        #endregion

        #region Costruttori

        public Downloader()
        {
            _basePath = ConfigurationManager.AppSettings["basePath"] ?? _basePath;
            _baseURL = ConfigurationManager.AppSettings["baseURL"] ?? _baseURL;
            _dwnldURL = ConfigurationManager.AppSettings["dwnldURL"] ?? _dwnldURL;

            if(!DateTime.TryParseExact(ConfigurationManager.AppSettings["data"], "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _data))
                _data = DateTime.Now;
        }

        #endregion

        #region Metodi

        public void Run()
        {
            try
            {
                _htmlDoc.LoadHtml(_webClient.DownloadString(_baseURL + _dwnldURL));

                //ottengo l'array delle date visualizzate
                HtmlNodeCollection nodes = _htmlDoc.DocumentNode.SelectNodes("//div[@class='DNN_Documents']//table//tr");

                foreach (var node in nodes)
                {
                    if (node.SelectSingleNode(".//td[@class='OwnerCell']") != null
                        && node.SelectSingleNode(".//td[@class='OwnerCell']").InnerText == "Previsione Produzione Eolica"
                        && node.SelectSingleNode(".//td[@class='CategoryCell']").InnerText == _data.ToString("dd/MM/yyyy"))
                    {
                        string link = node.SelectSingleNode(".//td[@class='OwnerCell']//a").Attributes["href"].Value;

                        Uri uri = new Uri(_baseURL + _dwnldURL);
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_baseURL + link);

                        request.Referer = uri.ToString();
                        request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
                        request.KeepAlive = true;
                        
                        //.Net 4.0
                        //request.Host = "www.terna.it";
                        
                        request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.90 Safari/537.36";
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        Stream stream = response.GetResponseStream();

                        using (var fileStream = File.Create(System.IO.Path.Combine(_basePath, "PrevProdEolica_" + _data.ToString("yyyyMMdd") + ".xls")))
                        {
                            byte[] buffer = new byte[16 * 1024]; // Fairly arbitrary size
                            int bytesRead;

                            while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                fileStream.Write(buffer, 0, bytesRead);
                            }

                            //.Net 4.0
                            //stream.CopyTo(fileStream);
                        }
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("ERRORE - {0}", e.Message);
            }
        }

        #endregion
    }
}
