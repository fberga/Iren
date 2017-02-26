using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Xml;
using System.Runtime.InteropServices;
using System.Collections;
using System.Configuration;


namespace Iren.FrontOffice.Grafici_IPEX_2014
{
    public partial class MainForm : Form
    {
        private const string _urlGrafici = @"http://www.mercatoelettrico.org/It/Esiti/MGP/DomandaOfferta.aspx?";
        private const string _urlMsd = @"http://www.mercatoelettrico.org/It/WebServerDataStore/MSD_ServiziDispacciamento/";
        //private const string _basePath = @"\\master.local\To\DATI3\MERCATO_ELETTRICO\Esiti\Grafici\";
        //private const string _basePath = @"C:\Users\e-bergamin\Desktop\Esiti\Grafici\";
        //private const string _basePathXML = @"\\srvebasbo.aem.torino.it\LocalDoc";

        private string _basePath;
        private string _basePathXML;

        [DllImport("wininet.dll", SetLastError = true)]
        public static extern bool InternetGetCookieEx(
            string url,
            string cookieName,
            StringBuilder cookieData,
            ref int size,
            Int32 dwFlags,
            IntPtr lpReserved);

        private const Int32 InternetCookieHttponly = 0x2000;

        public MainForm()
        {
            InitializeComponent();
            _basePath = ConfigurationManager.AppSettings["basePath"];
            _basePathXML = ConfigurationManager.AppSettings["XMLPath"];

        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            this.dtData.Value = DateTime.Now;
            cboZona.SelectedIndex = 0;

            txtOutputXML.Text = _basePathXML;
            dtData_ValueChanged(null, null);

            wbNaviga.Navigate(@"http://www.mercatoelettrico.org/It/Esiti/MGP/EsitiMGP.aspx");
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

        private void DownloadFileMSD()
        {
            if (txtOutputXML.Text == string.Empty)
            {
                MessageBox.Show("Selezionare le directory dove salvare i file.");
                return;
            }
            try
            {
                DateTime dt = dtData.Value.AddDays(-1);
                Uri uri = new Uri(_urlMsd + dt.ToString("yyyyMMdd") + "MSDServiziDispacciamento.xml");
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
                CookieContainer container = new CookieContainer();
                container.SetCookies(
                    uri,
                    GetGlobalCookies(uri.AbsoluteUri)
                );
                request.CookieContainer = container;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream stream = response.GetResponseStream();
                
                using (var fileStream = File.Create(System.IO.Path.Combine(txtOutputXML.Text, dt.ToString("yyyyMMdd") + "MSDServiziDispacciamento.xml")))
                {
                    stream.CopyTo(fileStream);
                }
            }
            catch (System.Net.WebException wex)
            {
                if (wex.Message.Contains("500"))
                {
                    DialogResult dr = MessageBox.Show("I file non sono stati scaricati, serve accettare le condizioni di GME. Premendo OK si verrà rediretti sulla pagina di accettazione", "Attenzione!", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                    if (dr == System.Windows.Forms.DialogResult.OK)
                        wbNaviga.Navigate(@"http://www.mercatoelettrico.org/It/Esiti/MGP/EsitiMGP.aspx");
                    return;
                }
                if (wex.Message.Contains("404"))
                {
                    MessageBox.Show("I file non sono stati scaricati perché non presenti nel server. Controllare più tardi.", "Attenzione!", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            

            MessageBox.Show("Salvataggio completato", "Download XML", MessageBoxButtons.OK);
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            btnLogin.Enabled = false;
            wbNaviga.Navigate(@"https://www.ipex.it");
        }
        private void btnOpen_Click(object sender, EventArgs e)
        {
            bdFolder.ShowDialog();
            txtOutput.Text = bdFolder.SelectedPath;
        }
        private void btnOpenXML_Click(object sender, EventArgs e)
        {
            bdFolder.ShowDialog();
            txtOutputXML.Text = bdFolder.SelectedPath;
        }

        private void btnSalva_Click(object sender, EventArgs e)
        {
            if (txtOutput.Text == string.Empty || txtOutputXML.Text == string.Empty)
            {
                MessageBox.Show("Selezionare le directory dove salvare i file.");
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            btnSalva.Enabled = false;
            loadingPanel.Visible = true;
            Application.DoEvents();

            bool is25hours = (dtData.Value.Month == 10 && isLastSunday(dtData.Value));
            bool is23hours = !is25hours && (dtData.Value.Month == 3 && isLastSunday(dtData.Value));

            int count = (is25hours ? 25 : (is23hours ? 23 : 24));

            if (chkTutte.Checked)
                cboZona.SelectedIndex = 0;

            progressBar1.Minimum = 1;
            progressBar1.Maximum = count - 1;
            progressBar1.Step = 1;
            do
            {
                progressBar1.Value = 1;
                for (int i = 1; i <= count; i++)
                {
                    string url = _urlGrafici;
                    url += "zona=" + cboZona.Text;
                    url += "&data=" + dtData.Value.ToString("yyyyMMdd");
                    url += "&ora=" + i.ToString();

                    Uri uri = new Uri(url);

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

                    CookieContainer container = new CookieContainer();
                    container.SetCookies(
                        uri,
                        GetGlobalCookies(uri.AbsoluteUri)
                    );

                    request.CookieContainer = container;
                    request.AllowAutoRedirect = false;
                    request.UserAgent = "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1300.0 Iron/23.0.1300.0 Safari/537.11";

                    try
                    {
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        Stream stream = response.GetResponseStream();
                        using (var fileStream = File.Create(System.IO.Path.Combine(txtOutput.Text, dtData.Value.ToString("yyyyMMdd") + cboZona.Text + i.ToString("00") + ".gif")))
                        {
                            stream.CopyTo(fileStream);
                        }
                    }
                    catch (System.Net.WebException wex)
                    {
                        DialogResult dr = MessageBox.Show("I file non sono stati scaricati, serve accettare le condizioni di GME. Premendo OK si verrà rediretti sulla pagina di accettazione", "Attenzione!", MessageBoxButtons.OKCancel);
                        if(dr == System.Windows.Forms.DialogResult.OK)
                            wbNaviga.Navigate(@"http://www.mercatoelettrico.org/It/Esiti/MGP/EsitiMGP.aspx");

                        return;
                    }
                    

                    lbFileName.Text = System.IO.Path.Combine(txtOutput.Text, dtData.Value.ToString("yyyyMMdd") + cboZona.Text + i.ToString("00") + ".gif");
                    progressBar1.PerformStep();
                    Application.DoEvents();
                }

                if (chkTutte.Checked && cboZona.SelectedIndex < cboZona.Items.Count - 1)
                    cboZona.SelectedIndex++;
                else
                    chkTutte.Checked = false;

            } while (chkTutte.Checked);

            lbFileName.Text = "";
            this.Cursor = Cursors.Arrow;
            btnSalva.Enabled = true;
            loadingPanel.Visible = false;
            DownloadFileMSD();
            
        }
        private void btnMSD_Click(object sender, EventArgs e)
        {
            DownloadFileMSD();
        }

        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InternetGetCookieEx(string pchURL, string pchCookieName, StringBuilder pchCookieData, ref uint pcchCookieData, int dwFlags, IntPtr lpReserved);
        const int INTERNET_COOKIE_HTTPONLY = 0x00002000;

        public static string GetGlobalCookies(string uri)
        {
            uint datasize = 1024;
            StringBuilder cookieData = new StringBuilder((int)datasize);
            if (InternetGetCookieEx(uri, null, cookieData, ref datasize, INTERNET_COOKIE_HTTPONLY, IntPtr.Zero)
                && cookieData.Length > 0)
            {
                return cookieData.ToString().Replace(';', ',');
            }
            else
            {
                return null;
            }
        }


        private void txtIdZone_TextChanged(object sender, EventArgs e)
        {
            btnSalva.Enabled = true;
        }
        private void txtOutput_TextChanged(object sender, EventArgs e)
        {
            btnSalva.Enabled = true;
        }
        private void dtData_ValueChanged(object sender, EventArgs e)
        {
            txtOutput.Text = _basePath + dtData.Value.ToString("yyyy") + @"\" + dtData.Value.ToString("MM") + @"\" + dtData.Value.ToString("dd");

            if (Directory.Exists(_basePath + dtData.Value.ToString("yyyy")) == false)
            {
                Directory.CreateDirectory(_basePath + dtData.Value.ToString("yyyy"));
            }

            if (Directory.Exists(_basePath + dtData.Value.ToString("yyyy") + @"\" + dtData.Value.ToString("MM")) == false)
            {
                Directory.CreateDirectory(_basePath + dtData.Value.ToString("yyyy") + @"\" + dtData.Value.ToString("MM"));
            }

            if (Directory.Exists(_basePath + dtData.Value.ToString("yyyy") + @"\" + dtData.Value.ToString("MM") + @"\" + dtData.Value.ToString("dd")) == false)
            {
                Directory.CreateDirectory(_basePath + dtData.Value.ToString("yyyy") + @"\" + dtData.Value.ToString("MM") + @"\" + dtData.Value.ToString("dd"));
            }
        }
    }
}