using Iren.PSO.Base;
using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Iren.PSO.Forms
{
    public partial class FormMeteo : Form
    {
        DataTable _entita;
        DataTable _entitaProprieta;
        DateTime _dataRif;
        ACarica _carica;
        ARiepilogo _riepilogo;

        public FormMeteo(object dataRif, ACarica carica, ARiepilogo riepilogo)
        {
            InitializeComponent();

            _carica = carica;
            _entita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA];
            _entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA];
            _dataRif = (DateTime)dataRif;
            _riepilogo = riepilogo;

            labelData.Text = "Data Riferimento:   " + _dataRif.ToString("dddd dd MMMM yyyy");
            this.Text = Simboli.NomeApplicazione + " - Meteo";
        }

        private void frmMETEO_Load(object sender, EventArgs e)
        {
            //_entitaProprieta.RowFilter = "SiglaProprieta = 'PROGR_IMPIANTO_TEMP_FONTE_ATTIVA' AND IdApplicazione = " + Workbook.IdApplicazione;

            var filter = _entitaProprieta
                .AsEnumerable()
                .Where(r => r["SiglaProprieta"].Equals("PROGR_IMPIANTO_TEMP_FONTE_ATTIVA") && r["IdApplicazione"].Equals(Workbook.IdApplicazione))
                .Select(r => "'" + r["SiglaEntita"] + "'")
                .ToList();

            string filtro = String.Join(",", filter);
            //foreach (DataRowView prop in _entitaProprieta)
            //{
            //    filtro += "'" + prop["SiglaEntita"] + "',";
            //}

            //if (filtro.Length > 0)
            //{
            //    filtro = "SiglaEntita IN (" + filtro + ")";
            //    _entita.RowFilter = filtro + " AND IdApplicazione = " + Workbook.IdApplicazione;
            //}
            DataView lista = new DataView(_entita);

            lista.RowFilter = "SiglaEntita IN (" + filtro + ") AND IdApplicazione = " + Workbook.IdApplicazione;

            comboUP.DisplayMember = "DesEntita";
            comboUP.ValueMember = "SiglaEntita";
            comboUP.DataSource = lista;
        }

        private void comboUP_SelectedIndexChanged(object sender, EventArgs e)
        {            
            if (DataBase.OpenConnection())
            {
                Array comboArray = groupDati.Controls.OfType<ComboBox>().ToArray();
                foreach (ComboBox cmb in comboArray)
                    groupDati.Controls.Remove(cmb);

                Array radioArray = groupDati.Controls.OfType<RadioButton>().ToArray();
                foreach (RadioButton rdb in radioArray)
                    groupDati.Controls.Remove(rdb);

                DataTable fonti = DataBase.Select(DataBase.SP.CHECK_FONTE_METEO, "@SiglaEntita=" + comboUP.SelectedValue + ";@Data=" + _dataRif.ToString("yyyyMMdd")) ?? new DataTable();

                int fonteOrdine = 0;
                foreach (DataRow fonte in fonti.Rows)
                {
                    DateTime dataEmissione = DateTime.ParseExact(fonte["DataEmissione"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                    if (!groupDati.Controls.ContainsKey("combo" + fonte["CodiceFonte"]))
                    {
                        ComboBox cmb = new ComboBox() 
                        { 
                            Name = "combo" + fonte["CodiceFonte"],
                            Font = groupDati.Font,
                            Location = new System.Drawing.Point(146, 50 + (28 * fonteOrdine) + 8),
                            Size = new System.Drawing.Size(190, 28),
                            FormattingEnabled = true
                        };
                        RadioButton rdb = new RadioButton()
                        {
                            Name = fonte["CodiceFonte"].ToString(),
                            Text = fonte["CodiceFonte"].ToString(),
                            Font = groupDati.Font,
                            Location = new System.Drawing.Point(5, 52 + (28 * fonteOrdine) + 8),
                            Size = new System.Drawing.Size(82, 24),
                            Checked = fonteOrdine == 0
                        };

                        rdb.CheckedChanged += rdb_CheckedChanged;

                        groupDati.Controls.Add(rdb);
                        groupDati.Controls.Add(cmb);
                        fonteOrdine++;
                    }
                    ((ComboBox)groupDati.Controls["combo" + fonte["CodiceFonte"]]).Items.Add(dataEmissione);
                    ((ComboBox)groupDati.Controls["combo" + fonte["CodiceFonte"]]).SelectedIndex = 0;
                }
            }
        }

        void rdb_CheckedChanged(object sender, EventArgs e)
        {
            if (DataBase.OpenConnection())
            {
                RadioButton rbt = (RadioButton)sender;

                if (rbt.Checked)
                {
                    //TODO eliminare questo filtro e passare direttamente il codice della fonte (DA AGGIORNARE STRUTTURA SU DB)
                    var proprieta = _entitaProprieta
                        .AsEnumerable()
                        .Where(r => r["SiglaProprieta"].Equals("PROGR_IMPIANTO_TEMP_FONTE") && r["SiglaEntita"].Equals(comboUP.SelectedValue) && r["Valore"].Equals(rbt.Name) && r["IdApplicazione"].Equals(Workbook.IdApplicazione))
                        .Select(r => r["Ordine"])
                        .FirstOrDefault();

                    DataBase.Insert(DataBase.SP.UPDATE_METEO, new Core.QryParams() 
                    {
                        {"@SiglaEntita",comboUP.SelectedValue},
                        {"@Valore", proprieta}
                    });
                }
            }
            
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            if (DataBase.OpenConnection())
            {
                //TODO passare direttamente il codice della fonte (DA AGGIORNARE STRUTTURA SU DB)
                foreach (DataRow entita in _entita.Rows)
                {
                    DataBase.Insert(DataBase.SP.UPDATE_METEO, new Core.QryParams() 
                        {
                            {"@SiglaEntita", entita["SiglaEntita"]},
                            {"@Valore", "1"}
                        });
                }
                this.Close();
            }
        }

        private void btnCarica_Click(object sender, EventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            btnCarica.Enabled = false;
            btnAnnulla.Enabled = false;

            //object siglaEntita = comboUP.SelectedValue;

            string nomeCombo = "combo" + groupDati.Controls.OfType<RadioButton>().FirstOrDefault(btn => btn.Checked).Name;
            ComboBox cmb = groupDati.Controls.OfType<ComboBox>().FirstOrDefault(c => c.Name == nomeCombo);

            string dataEmissione = ((DateTime)cmb.SelectedItem).ToString("yyyyMMdd");

            bool gone = _carica.AzioneInformazione(comboUP.SelectedValue, "METEO", "CARICA", _dataRif, null, dataEmissione);

            _riepilogo.AggiornaRiepilogo(comboUP.SelectedValue, "METEO", gone, _dataRif);

            Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogCarica, "Carica: Previsioni meteo");

            btnCarica.Enabled = true;
            btnAnnulla.Enabled = true;

            Workbook.ScreenUpdating = true;
            Sheet.Protected = true;
        }
    }
}
