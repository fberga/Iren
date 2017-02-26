using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GeneraXls
{
    public partial class LoadForm : Form
    {
        #region Variables

        /// <summary>
        /// DataTable which displays available UPs for the generation.
        /// </summary>
        private DataTable _dtCentrali = new DataTable("dtCentrali") 
        {
            Columns =
            {
                new DataColumn("Centrale", typeof(string)),
                new DataColumn("XML", typeof(string)),
                new DataColumn("pathXML", typeof(string)),
                new DataColumn("CSV", typeof(string)),
                new DataColumn("pathCSV", typeof(string))
            }
        };

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor.
        /// </summary>
        public LoadForm()
        {
            InitializeComponent();
            //set the datasource
            dataGridCentrali.DataSource = _dtCentrali;
            dataGridCentrali.Columns["pathXML"].Visible = false;
            dataGridCentrali.Columns["pathCSV"].Visible = false;
            //define columns' dimensions
            dataGridCentrali.Columns["Centrale"].Width = dataGridCentrali.Width / 3;
            dataGridCentrali.Columns["XML"].Width = dataGridCentrali.Width / 6;
            dataGridCentrali.Columns["CSV"].Width = dataGridCentrali.Width / 6;

            //disable execute button
            btnGenera.Enabled = false;

            //find all the available UPs in the App.config
            string[] ups =
                ConfigurationManager.AppSettings.AllKeys
                    .Where(f => Regex.IsMatch(f, @"cod_rup\d+"))
                    .ToArray();

            //initialize the DataTable with default values
            foreach (string up in ups)
            {
                DataRow r = _dtCentrali.NewRow();
                r["Centrale"] = ConfigurationManager.AppSettings[up];
                r["XML"] = r["CSV"] = "no";
                _dtCentrali.Rows.Add(r);
            }

            //set input/output default paths from App.config
            txtPathInput.Text = ConfigurationManager.AppSettings["pathInput"];
            txtPathOutput.Text = ConfigurationManager.AppSettings["pathOutput"];

            //retireve all defined MSD markets from App.config
            cmbMercato.Items.AddRange(ConfigurationManager.AppSettings["markets"].Split(new char[] { ',', ';' }));
            cmbMercato.SelectedIndex = 0;

            //set default data to Today
            dtpData.Value = DateTime.Today;

            //enable events on textChanged
            txtPathInput.TextChanged += CheckFilesInputDirectory;
            dtpData.TextChanged += CheckFilesInputDirectory;
            cmbMercato.TextChanged += CheckFilesInputDirectory;

            //enable events of choose path
            chooseFolder.ShowNewFolderButton = false;
            btnSfogliaPathInput.Click += ChoosePath;
            btnSfogliaPathOutput.Click += ChoosePath;

            //enable generation event
            btnGenera.Click += Genera;

            //trigger first check at opening
            CheckFilesInputDirectory(null, null);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Starting from the App.config, it retrieves all path formats and replaces the variables with the correct values.
        /// </summary>
        /// <param name="date">Date of the file.</param>
        /// <param name="codRup">RUP code of the file.</param>
        /// <param name="msd">Market of the file.</param>
        /// <returns>The list of available paths for the combination of the input variables.</returns>
        private string[] GetPath(DateTime date, string codRup, string msd)
        {
            //get all "pathFormats"
            string[] keys =
                ConfigurationManager.AppSettings.AllKeys
                    .Where(f => Regex.IsMatch(f, @"pathFormat\d+"))
                    .ToArray();

            string[] o = new string[keys.Length];
            int i = 0;
            foreach (string k in keys)
            {
                string format = ConfigurationManager.AppSettings[k];
                string dateFormat = Regex.Match(format, @"\[DATE_[yMd]*\]").Value;
                dateFormat = dateFormat.Replace("[DATE_", "").Replace("]", "");

                o[i] = Regex.Replace(format, @"\[DATE_[yMd]*\]", date.ToString(dateFormat), RegexOptions.IgnoreCase);
                o[i] = Regex.Replace(o[i], @"\[MSD\]", msd, RegexOptions.IgnoreCase);
                o[i] = Regex.Replace(o[i++], @"\[CODRUP\]", codRup, RegexOptions.IgnoreCase);
            }

            return o;
        }

        /// <summary>
        /// Parse XML files.
        /// </summary>
        /// <param name="fileNameXml">path of the file.</param>
        /// <param name="wb">Workbook.</param>
        /// <param name="codRup">Code RUP.</param>
        /// <returns>true if the parsing completed, false otherwise.</returns>
        private bool ReadXml(string fileNameXml, Excel.Workbook wb, string codRup)
        {
            //load xml
            XDocument xmldoc = XDocument.Load(fileNameXml);

            //get default namespace
            XNamespace ns = xmldoc.Root.GetDefaultNamespace();
            var hourDetails = xmldoc.Descendants(ns + "HourDetail");

            //first cell of the range
            int row = -1;
            int col = -1;
            try
            {
                Excel.Range rng = wb.Sheets[1].Range[codRup];

                row = rng.Row;
                col = rng.Column;

                Marshal.ReleaseComObject(rng);
                rng = null;

            }
            catch
            {
                MessageBox.Show("Si sono verificati errori nella ricerca della centrale " + codRup + " nel file excel. Contattare l'amministratore.", "ERRORE", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }
            
            try 
            {
                foreach (XElement hourDetail in hourDetails)
                {
                    int hour = int.Parse(hourDetail.Element(ns + "Hour").Value) - 1;

                    foreach (XElement quantity in hourDetail.Elements(ns + "Quantity"))
                    {
                        int quarter = int.Parse(quantity.Attribute("QuarterInterval").Value) - 1;
                        
                        /************************ Modifica calcolo valori potenza da xml ********************/
                        /************************           02/02/2017    BEGIN          ********************/
                        /*
                        wb.Sheets[1].Cells[row + hour, col + quarter].Value = (Double.Parse((quantity.Value.Replace('.', ','))));
                         */
                        Double val = Double.Parse((quantity.Value.Replace('.', ',')));
                        val = val*4;
                        wb.Sheets[1].Cells[row + hour, col + quarter].Value = val;
                        /************************ Modifica calcolo valori potenza da xml ********************/
                        /************************            02/02/2017    END           ********************/
                    }
                }
            }
            catch
            {
                MessageBox.Show("Si sono verificati errori nella lettura dell'XML di " + codRup + ". Verificare il file e rilanciare la generazione.", "ERRORE", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }

            return true;
        }

        /// <summary>
        /// Parse CSV files.
        /// </summary>
        /// <param name="fileNameCSV"> Path of the file.</param>
        /// <param name="wb">Workbook.</param>
        /// <param name="codRup">Code RUP.</param>
        /// <returns></returns>
        private bool ReadCSV(string fileNameCSV, Excel.Workbook wb, string codRup)
        {
            string[] lines = System.IO.File.ReadAllLines(fileNameCSV);

            //first cell of the range
            int row = -1;
            int col = -1;
            try
            {
                Excel.Range rng = wb.Sheets[1].Range[codRup];

                row = rng.Row;
                col = rng.Column;

                Marshal.ReleaseComObject(rng);
                rng = null;

            }
            catch
            {
                MessageBox.Show("Si sono verificati errori nella ricerca della centrale " + codRup + " nel file excel. Contattare l'amministratore.", "ERRORE", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }

            try
            {
                foreach (string line in lines)
                {
                    string[] cols = line.Split(new char[] { ';', ' ' });
                    int hour = int.Parse(cols[3]) - 1;
                    wb.Sheets[1].Cells[row + hour, col] = cols[4];
                    wb.Sheets[1].Cells[row + hour, col + 1] = cols[5];
                    wb.Sheets[1].Cells[row + hour, col + 2] = cols[6];
                    wb.Sheets[1].Cells[row + hour, col + 3] = cols[7];
                }
            }
            catch
            {
                MessageBox.Show("Si sono verificati errori nella lettura del CSV di " + codRup + ". Verificare il file e rilanciare la generazione.", "ERRORE", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }

            return true;
        }

        /// <summary>
        /// Get total hours of the input day.
        /// </summary>
        /// <param name="giorno">Input day.</param>
        /// <returns>Total hours of the input day.</returns>
        private int GetOreGiorno(DateTime giorno)
        {
            DateTime giornoSucc = giorno.AddDays(1);
            return (int)(giornoSucc.ToUniversalTime() - giorno.ToUniversalTime()).TotalHours;
        }

        #endregion

        #region Events

        /// <summary>
        /// Generate the output excel.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Genera(object sender, EventArgs e)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            try
            {
                btnGenera.Enabled = false;
                btnChiudi.Enabled = false;
                //get the number of hours for selected date
                int oreGG = GetOreGiorno(this.dtpData.Value);

                //template path
                string pathTemplate = System.IO.Path.Combine(Environment.CurrentDirectory, @"template", "Template" + oreGG + ".xlt");

                //output name
                string fileNameXls = this.cmbMercato.SelectedItem + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";

                xlApp = new Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!");
                    return;
                }

                //open template
                wb = xlApp.Workbooks.Open(pathTemplate);

                //write header data
                wb.Sheets[1].Range["C1"] = this.dtpData.Value;
                wb.Sheets[1].Range["F1"] = this.cmbMercato.SelectedItem;

                bool finished = true;
                foreach (DataRow r in _dtCentrali.Rows)
                {
                    //execute parsing for available extensions
                    if (r["XML"].Equals("sì"))
                    {
                        if (!ReadXml(r["pathXML"].ToString(), wb, r["Centrale"].ToString()))
                        {
                            finished = false;
                            break;
                        }

                    }
                    else if (r["CSV"].Equals("sì"))
                    {
                        if (!ReadCSV(r["pathCSV"].ToString(), wb, r["Centrale"].ToString()))
                        {
                            finished = false;
                            break;
                        }
                    }
                }

                //if everything worked, save a copy of the template
                if (finished)
                    wb.SaveAs(Path.Combine(txtPathOutput.Text, fileNameXls), Excel.XlFileFormat.xlExcel8);

                //close and clean
                wb.Close(SaveChanges: false);

                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(xlApp);

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Genera XLS - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                wb = null;
                xlApp = null;
                btnGenera.Enabled = true;
                btnChiudi.Enabled = true;
            }
        }

        /// <summary>
        /// Check weather the files are available or not. If multiple file are available for the same UP,
        /// the first for each extension is considered. Enable Execute button when all requirements are satisfied.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckFilesInputDirectory(object sender, EventArgs e)
        {
            foreach (DataRow r in _dtCentrali.Rows)
            {
                //get all available formats
                string[] files =
                    GetPath(dtpData.Value,
                        r["Centrale"].ToString(),
                        cmbMercato.Items[cmbMercato.SelectedIndex].ToString());

                r["XML"] = "no";
                r["CSV"] = "no";
                r["pathXML"] = "";
                r["pathCSV"] = "";

                foreach (string file in files)
                {
                    //complete path
                    string path = System.IO.Path.Combine(txtPathInput.Text, file);

                    //check existence
                    if (System.IO.File.Exists(path))
                    {
                        //keep only first file
                        string ext = System.IO.Path.GetExtension(path).Replace(".", "").ToUpper();
                        if (!r[ext].Equals("sì"))
                        {
                            r[ext] = "sì";
                            r["path" + ext] = path;
                        }
                    }
                }
            }

            bool isGeneraAvailable = true;
            foreach (DataRow r in _dtCentrali.Rows)
            {
                isGeneraAvailable &= r["XML"].Equals("sì") || r["CSV"].Equals("sì");
            }

            btnGenera.Enabled = isGeneraAvailable;
        }

        /// <summary>
        /// Close the application.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChiudi_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Allow user to browse the filesystem.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChoosePath(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            TextBox txt = null;

            switch (btn.Name)
            {
                case "btnSfogliaPathInput":
                    txt = txtPathInput;
                    break;
                case "btnSfogliaPathOutput":
                    txt = txtPathOutput;
                    break;
            }

            chooseFolder.SelectedPath = txt.Text;

            if (chooseFolder.ShowDialog() == DialogResult.OK)
            {
                txt.Text = chooseFolder.SelectedPath;
            }
        }

        #endregion

    }
}
