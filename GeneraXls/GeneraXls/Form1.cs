using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection; // for the Missing.Value

namespace GeneraXls
{
    public partial class FrmGeneraXls : Form
    {
        #region Variabili

        // the form parameters
        string tipoMercato;
        string[] listaMercati;
        string pathInput;
        string pathOutput;

        //string nomeXmlRosone;
        //string nomeXmlOrco;
        //string nomeXmlTelessio;
        //string nomeXmlVilla;
        List<string> _filesToRead;
        
        #endregion



        #region Costruttore

        public FrmGeneraXls()
        {
            InitializeComponent();
            // read parameters from App.config
            tipoMercato = ConfigurationManager.AppSettings.Get("tipoMercato");
            pathInput = ConfigurationManager.AppSettings.Get("pathInput");
            pathOutput = ConfigurationManager.AppSettings.Get("pathOutput");

            char[] charSeparators = new char[] {',',' '};
            listaMercati = tipoMercato.Split(charSeparators, StringSplitOptions.None);
            this.cmbMercato.Items.Add(listaMercati[0]);
            this.cmbMercato.Items.Add(listaMercati[1]);
            this.cmbMercato.Items.Add(listaMercati[2]);
            this.cmbMercato.Items.Add(listaMercati[3]);
            this.cmbMercato.SelectedItem = listaMercati[0];

            this.txtPathInput.SelectedText = pathInput;
            this.txtPathOutput.Text = pathOutput;

            // the "Genera" button is disabled at the start
            this.btnGenera.Enabled = false;
            // 
            SetLabelXmlRed();
            
            // ATTENZIONE: se inserisco queste due righe si apre due volte la dialog
            // registro i gestori degli eventi
            //this.btnSfogliaPathInput.Click += new System.EventHandler(this.btnSfogliaPathInput_Click);
            //this.btnSfogliaPathOutput.Click += new System.EventHandler(this.btnSfogliaPathOutput_Click);
        }

        #endregion

        

        #region Metodi

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
        /// Setta le labels dell'elenco dei file xml in Rosso
        /// </summary>
        public void SetLabelXmlRed()
        {
            this.lblXmlRosone.ForeColor = System.Drawing.Color.Red;
            this.lblXmlOrco.ForeColor = System.Drawing.Color.Red;
            this.lblXmlTelessio.ForeColor = System.Drawing.Color.Red;
            this.lblXmlVilla.ForeColor = System.Drawing.Color.Red;
        }

        /// <summary>
        /// Lettura XML e scrittura XLS
        /// </summary>
        /// <param name="fileNameXml">Path file XML.</param>
        /// <param name="wb">Workbook.</param>
        /// <param name="columns">Colonne XLS.</param>
        private void ReadXml(string fileNameXml, Excel.Workbook wb, string[] columns)
        {
            XmlDocument xmldoc = new XmlDocument();
            XmlNodeList xmlnode;
            FileStream fs = new FileStream(fileNameXml, FileMode.Open, FileAccess.Read);
            xmldoc.Load(fs);
            xmlnode = xmldoc.GetElementsByTagName("HourDetail");
            string row = null;
            if (xmlnode.Count > 0)
            {
                for (int i = 0; i <= xmlnode.Count - 1; i++)
                {
                    row = (i + 5).ToString();
                    // i Quarto D'ora
                    wb.Sheets[1].Range[columns[0] + row] = Double.Parse((xmlnode[i].ChildNodes.Item(1).InnerText.Trim()).ToString().Replace(".", ","));
                    // ii Quarto D'ora
                    wb.Sheets[1].Range[columns[1] + row] = Double.Parse((xmlnode[i].ChildNodes.Item(2).InnerText.Trim()).ToString().Replace(".", ","));
                    // iii Quarto D'ora
                    wb.Sheets[1].Range[columns[2] + row] = Double.Parse((xmlnode[i].ChildNodes.Item(3).InnerText.Trim()).ToString().Replace(".", ","));
                    // iv Quarto D'ora
                    wb.Sheets[1].Range[columns[3] + row] = Double.Parse((xmlnode[i].ChildNodes.Item(4).InnerText.Trim()).ToString().Replace(".", ","));
                }
            }
            else
            {
                MessageBox.Show("Il file XML non è ben strutturato.");
            }
            xmldoc = null;
            xmlnode = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void ReadCSV(string fileNameCSV, Excel.Workbook wb, string[] columns)
        {
            string[] lines = System.IO.File.ReadAllLines(fileNameCSV);

            int row = 4;
            foreach (string line in lines)
            {
                string[] cols = line.Split(';');
                int hour = int.Parse(cols[3]) + row;
                wb.Sheets[1].Range[columns[0] + hour] = cols[4];
                wb.Sheets[1].Range[columns[1] + hour] = cols[5];
                wb.Sheets[1].Range[columns[2] + hour] = cols[6];
                wb.Sheets[1].Range[columns[3] + hour] = cols[7];
            }
        }

        #endregion


        #region Eventi

        private void btnSfogliaPathInput_Click(object sender, EventArgs e)
        {
            SetLabelXmlRed();

            folderBrowserDialogInput.SelectedPath = this.txtPathInput.Text;
            if (folderBrowserDialogInput.ShowDialog() == DialogResult.OK)
            {
                this.txtPathInput.Clear();
                this.txtPathInput.Text = folderBrowserDialogInput.SelectedPath;
            }

            /*
             * qui cerco se nella cartella di INPUT ho TUTTI i 4 files XML
             */
            string endOfPath = this.cmbMercato.SelectedItem + "D_" + this.dtpData.Value.ToString("yyyyMMdd") + ".xml.OEIESRD.out.xml";
            //nomeXmlRosone = this.txtPathInput.Text + @"\FMS_UP_ROSONE_1_" + endOfPath;
            //nomeXmlOrco = this.txtPathInput.Text + @"\FMS_UP_ORCO_1_" + endOfPath;
            //nomeXmlTelessio = this.txtPathInput.Text + @"\FMS_UP_TELESSIO_1_" + endOfPath;
            //nomeXmlVilla = this.txtPathInput.Text + @"\FMS_UP_VILLA_1_" + endOfPath;

            // search all file xml for "data" and "mercato"
            _filesToRead = Directory.GetFiles(this.txtPathInput.Text, "FMS_UP_*" + endOfPath, SearchOption.TopDirectoryOnly).ToList<string>();

            bool isOrcoXMLAvailable = _filesToRead.Where(s => s.Contains("ORCO")).Count() > 0;
            if(!isOrcoXMLAvailable) 
            {
                //prompt to choose csv
                if(MessageBox.Show("Attenzione, il file xml di ORCO non è presente nella cartella. Selezionare un altro file?", "ATTENZIONE", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    OpenFileDialog ofd = new OpenFileDialog();
                    ofd.InitialDirectory = this.txtPathInput.Text;
                    ofd.Filter = "*.xml|*.csv";

                    if(ofd.ShowDialog() == DialogResult.OK) 
                    {
                        _filesToRead.Add(ofd.FileName);
                    }
                }
            }

            int count = 0;
            foreach (string file in _filesToRead)
            {
                if (file.Contains("ROSONE"))
                {
                    this.lblXmlRosone.ForeColor = System.Drawing.Color.Green;
                    count++;
                }
                else if (file.Contains("ORCO"))
                {
                    this.lblXmlOrco.ForeColor = System.Drawing.Color.Green;
                    count++;
                }
                else if (file.Contains("TELESSIO"))
                {
                    this.lblXmlTelessio.ForeColor = System.Drawing.Color.Green;
                    count++;
                }
                else if (file.Contains("VILLA"))
                {
                    this.lblXmlVilla.ForeColor = System.Drawing.Color.Green;
                    count++;
                }
            }
            if (count == 4)
            {
                this.btnGenera.Enabled = true;
            }
            else
            {
                MessageBox.Show("Si è verificato un errore durante il processo: non trovo il file per una centrale.");
            }

            folderBrowserDialogInput.Dispose();
        }

        private void btnSfogliaPathOutput_Click(object sender, EventArgs e)
        {
            folderBrowserDialogOutput.SelectedPath = this.txtPathOutput.Text;
            if (folderBrowserDialogOutput.ShowDialog() == DialogResult.OK)
            {
                this.txtPathOutput.Clear();
                this.txtPathOutput.Text = folderBrowserDialogOutput.SelectedPath;
            }
            folderBrowserDialogOutput.Dispose();
        }


        private void btnGenera_Click(object sender, EventArgs e)
        {
            // path and name of file excel
            string fileNameXls = @"\" + this.cmbMercato.SelectedItem + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            string fileNameCSV = @"\" + this.cmbMercato.SelectedItem + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
            string fileNameXlsFull = this.txtPathOutput.Text + fileNameXls;

            // we determine which model to use
            int oreGG = GetOreGiorno(this.dtpData.Value);
            string nameTemplate = @"Template" + oreGG + ".xlt";
            string pathTemplate = Path.Combine(Environment.CurrentDirectory, @"template\", nameTemplate);

            // initialize the Excel application Object
            Excel.Application xlApp = new Excel.Application();
            // check if Excel is installed in your system
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!");
                return;
            }

            Excel.Workbook wb = xlApp.Workbooks.Open(pathTemplate);

            wb.Sheets[1].Range["C1"] = this.dtpData.Value;
            wb.Sheets[1].Range["F1"] = this.cmbMercato.SelectedItem;

            if (_filesToRead.Count > 0)
            {
                foreach (string file in _filesToRead)
                {
                    if (file.Contains("ROSONE"))
                    {
                        ReadXml(file, wb, new string[] { "C", "D", "E", "F" });
                    }
                    else if (file.Contains("ORCO"))
                    {
                        try
                        {
                            ReadXml(file, wb, new string[] { "H", "I", "J", "K" });
                        }
                        catch
                        {
                            ReadCSV(file, wb, new string[] { "H", "I", "J", "K" });
                        }
                        
                    }
                    else if (file.Contains("TELESSIO"))
                    {
                        ReadXml(file, wb, new string[] { "M", "N", "O", "P" });
                    }
                    else if (file.Contains("VILLA"))
                    {
                        ReadXml(file, wb, new string[] { "R", "S", "T", "U" });
                    }
                }
                MessageBox.Show("Generazione del file excel andata a buon fine.");
                this.btnGenera.Enabled = false;
            }
            else
            {
                MessageBox.Show("Si è verificato un errore durante il processo: Non ho trovato i files XML");
            }

            //save export XLS and close
            wb.SaveAs(fileNameXlsFull, Excel.XlFileFormat.xlExcel8);
            wb.Close();

            Marshal.ReleaseComObject(wb);
        }

        private void dtpData_ValueChanged(object sender, EventArgs e)
        {
            this.btnGenera.Enabled = false;
            SetLabelXmlRed();
        }

        private void cmbMercato_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.btnGenera.Enabled = false;
            SetLabelXmlRed();
        }

        #endregion

        private void btnChiudi_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
