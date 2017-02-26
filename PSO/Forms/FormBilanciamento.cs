using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Forms
{
    public partial class FormBilanciamento : Form
    {
        public const string MODIFICA = "FormIncrementoModifica";
        public const string ALL_OLD_VALUE = "FormIncrementoRipristino";

        private DefinedNames _definedNames;
        private Excel.Worksheet _ws;

        private Excel.Range _origRng;

        private DataTable _dtEntitaGradinoDisp;

        public FormBilanciamento(Excel.Worksheet ws, Excel.Range rng)
        {
            InitializeComponent();
            this.Text = Simboli.NomeApplicazione + " - Bilanciamento";
            
            _ws = ws;

            _ws.SelectionChange += ChangeBalanceHours;
            _definedNames = new DefinedNames(_ws.Name, DefinedNames.InitType.All);

            Workbook.Repository.Add(Workbook.Repository.CreaTabellaModifica(MODIFICA));
            Workbook.Repository.Add(Workbook.Repository.CreaTabellaRipristinaIncremento(ALL_OLD_VALUE));

            if (!CheckPrecondition())
            {
                //MessageBox.Show("Non sono state selezionate celle valide. Selezionare celle offerta energia.");
                return;
            }

            
            ChangeBalanceHours(Workbook.Application.Selection);

            
        }

        private void ChangeBalanceHours(Excel.Range Target)
        {
            if (!CheckPrecondition())
            {
                dgvBilanciamento.DataSource = null;
                return;
            }

            dgvBilanciamento.DataSource = null;
            _dtEntitaGradinoDisp = new DataTable()
            {
                Columns = 
                {
                    {"SiglaEntita", typeof(string)}
                }
            };

            _dtEntitaGradinoDisp.PrimaryKey = new DataColumn[] { _dtEntitaGradinoDisp.Columns["SiglaEntita"] };

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            DataView entitaInformazione = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
            DataView categoria = Workbook.Repository[DataBase.TAB.CATEGORIA].DefaultView;

            categoria.RowFilter = "DesCategoria = '" + _ws.Name + "'";

            categoriaEntita.RowFilter = "SiglaEntita <> '" + lbSiglaEntitaPrinc.Tag + "' AND Zona = '" + lbZonaUPprinc.Text + "' AND IdApplicazione = " + Workbook.IdApplicazione + " AND SiglaCategoria = '" + categoria[0]["SiglaCategoria"] + "' AND Gerarchia is null";

            //dgvBilanciamento.DataSource = categoriaEntita;

            int hCount = Target.Columns.Count;

            DataTable dt = new DataTable()
            {
                Columns =
                {
                    {"Priorità", typeof(string)},
                    {"DesEntita", typeof(string)},
                    {"SiglaEntita", typeof(string)}
                }
            };
            
            foreach (Excel.Range h in Target)
            {
                dt.Columns.Add(new DataColumn(""+h.Column, typeof(double)));
                _dtEntitaGradinoDisp.Columns.Add(new DataColumn("" + h.Column, typeof(string)));
            }
            

            string siglaInfoEnergia = GetInfoFromRange(Target);
            string siglaInfoTipo = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 1) + "TIPO";
            
            //cerco acq - ven
            //int priorityvalue = categoriaEntita.Count;
            foreach (DataRowView entita in categoriaEntita)
            {
                DataRow offSecondaria = dt.NewRow();
                int row = _definedNames.GetRowByName(lbSiglaEntitaPrinc.Tag, siglaInfoTipo);
                DataRow gradiniDisp = _dtEntitaGradinoDisp.NewRow();

                gradiniDisp["SiglaEntita"] = entita["SiglaEntita"];

                //offSecondaria["Priorità"] = priorityvalue;
                offSecondaria["SiglaEntita"] = entita["SiglaEntita"];
                offSecondaria["DesEntita"] = entita["DesEntita"];

                foreach (Excel.Range h in Target)
                {
                    string siglaInfo = _ws.Cells[row, h.Column].Value.Equals("VEN") ? "TEMP_OMI11" : "TEMP_OMI10";

                    int row1 = _definedNames.GetRowByName(entita["SiglaEntita"], siglaInfo);

                    offSecondaria["" + h.Column] = _ws.Cells[row1, h.Column].Value2;
                    for (int j = 1; j < 5; j++)
                    {
                        string siglaGradino = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 2) + j + "E";
                        int row2 = _definedNames.GetRowByName(entita["SiglaEntita"], siglaGradino);
                        if ( _ws.Cells[row2, h.Column].Value2 == null)
                        {
                            gradiniDisp["" + h.Column] = siglaGradino;
                            break;
                        }
                            
                    }
                }
                _dtEntitaGradinoDisp.Rows.Add(gradiniDisp);
                dt.Rows.Add(offSecondaria);
            }

            dgvBilanciamento.DataSource = dt;
            dgvBilanciamento.Columns["DesEntita"].ReadOnly = true;
            dgvBilanciamento.Columns["SiglaEntita"].Visible = false;
            foreach (Excel.Range h in Target)
            {
                dgvBilanciamento.Columns["" + h.Column].HeaderText = Date.GetSuffissoOra(h.Column - _definedNames.GetFirstCol());
                dgvBilanciamento.Columns["" + h.Column].ReadOnly = true;
            }

            _origRng = Target;
        }

        private bool CheckPrecondition()
        {
            //precondizione: Sono in un foglio di categoria
            /*if (!Workbook.CategorySheets.Contains(Workbook.ActiveSheet))
            {
                labelError.Text = "Nel foglio selezionato non è possibile eseguire questa operazione.";
                return false;
            }*/

            Excel.Range Target = Workbook.Application.Selection;

            if (Target.Rows.Count > 1)
            {
                foreach (Excel.Range row in Target.Rows)
                {
                    if (row.EntireRow.Hidden)
                    {
                        labelError.Text = "ERRORE: Nel range selezionato ci sono righe nascoste.";
                        return false;
                    }
                }

                labelError.Text = "ATTENZIONE: Nel range selezionato ci sono più righe.";
                return false;
            }

            foreach (Excel.Range row in Target.Rows)
            {
                if (!_definedNames.IsEditable(row.Row))
                {
                    labelError.Text = "ERRORE: Il range selezionato contiene delle righe non modificabili.";
                    return false;
                }
            }

            foreach (Excel.Range cell in Target)
            {
                if (cell.Value2 == null || cell.Value2.ToString() == "")
                {
                    labelError.Text = "ERRORE: Il range selezionato contiene delle righe non valorizzate. Per procedere al bilanciamento selezionare celle con valori maggiori di zero.";
                    return false;
                }
            }


            //precondizione: UP selezionata
            //_ws = Workbook.ActiveSheet;

            DataRow entita = GetEntitaFromRange(Target);

            if(entita == null)
                return false;

            //precondizione: Selezionata riga energia
            string siglaInformazione = GetInfoFromRange(Workbook.Application.Selection);

            if (!Regex.IsMatch(siglaInformazione, @"OFFERTA_MI\d_G\dE"))
                return false;

            //imposto valori da visualizzare
            lbSiglaEntitaPrinc.Text = entita["DesEntita"].ToString();
            lbSiglaEntitaPrinc.Tag = entita["SiglaEntita"];

            lbZonaUPprinc.Text = entita["Zona"].ToString();

            labelError.Text = "";
            return true;
        }

        private void ElementsEnabled(bool enabled)
        {
            /*cmbUpDaBilanciare.Enabled = enabled;
            cmbZone.Enabled = enabled;*/

            btnModifica.Enabled = enabled;
            btnRipristina.Enabled = enabled;

            dgvBilanciamento.Enabled = enabled;
        }

        private void Btn_Esegui_Click(object sender, EventArgs e)
        {
            DataView dv = ((DataTable)dgvBilanciamento.DataSource).AsDataView();
            Dictionary<string, double> traceBalance = new Dictionary<string, double>();

            dv.RowFilter = " Priorità is not null ";
            dv.Sort = "Priorità ASC";

            int column;

            string siglaInfoEnergia = GetInfoFromRange(_origRng);
            string siglaInfoTipo = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 1) + "TIPO";
            int row_tipo = _definedNames.GetRowByName(lbSiglaEntitaPrinc.Tag, siglaInfoTipo);

            int time = (int)DateTime.Now.TimeOfDay.TotalSeconds;

            foreach (Excel.Range cell in _origRng )
            {
                column = cell.Column;

                string op_sec_type = "";
                if (_ws.Cells[row_tipo, column].Value2.Equals("VEN"))
                {
                    op_sec_type = "ACQ";
                }
                else
                    op_sec_type = "VEN";

                double residuo = _ws.Cells[cell.Row, column].Value2;

                // Calcolo i contributi
                foreach (DataRowView dtv in dv)
                {
                    double secondaryOffert;
                    double.TryParse(dtv["" + column].ToString(), out secondaryOffert);
                    if (residuo >= secondaryOffert)
                    {
                        residuo -= secondaryOffert;

                        // salvo gli UP coinvolti nel bilanciamento
                        traceBalance.Add(dtv["SiglaEntita"].ToString(), secondaryOffert);
                        if (residuo == 0)
                            break;
                    }
                    else
                    {
                        traceBalance.Add(dtv["SiglaEntita"].ToString(), residuo);
                        residuo = 0;
                        break;
                    }
                }

                // Aggiorno il foglio con i contributi per l'ora esaminata

                Sheet.Protected = false;
                int column_con = column - _definedNames.GetFirstCol();
                string codBalance = (column_con < 10 ? "0" + column_con : "" + column_con) + time;

                //scrivo il codice bilanciamento nell'UP principale
                string sigla_balance = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 1) + "CB";
                int row_balance = _definedNames.GetRowByName(lbSiglaEntitaPrinc.Tag, sigla_balance);
                Handler.SaveOriginValues(_ws.Cells[row_balance, column], tableName: ALL_OLD_VALUE);
                _ws.Cells[row_balance, column].Value2 = codBalance;
                Handler.StoreEdit(_ws.Cells[row_balance, column], tableName: MODIFICA);

                // ciclo per tutti gli Up coinvolti nel bilanciamento
                foreach (string k in traceBalance.Keys)
                {
                    string siglaGradino = _dtEntitaGradinoDisp.Rows.Find(new object[] { k })["" + column].ToString();
                    int tmp_row = _definedNames.GetRowByName(k, siglaGradino);
                    Handler.SaveOriginValues(_ws.Cells[tmp_row, column], tableName: ALL_OLD_VALUE);
                    _ws.Cells[tmp_row, column].Value2 = traceBalance[k];
                    Handler.StoreEdit(_ws.Cells[tmp_row, column], tableName: MODIFICA);
                    // modifico il prezzo e lo imposto = 0
                    Handler.SaveOriginValues(_ws.Cells[tmp_row + 1, column], tableName: ALL_OLD_VALUE);
                    _ws.Cells[tmp_row+1, column].Value2 = 0; // campbiare secondo le proprietà dei nomi cella
                    Handler.StoreEdit(_ws.Cells[tmp_row + 1, column], tableName: MODIFICA);

                    // modifico la label ACQ/VEN
                    string type_op = siglaGradino.Substring(0, siglaGradino.Length - 1) + "TIPO";
                    int row_type_op = _definedNames.GetRowByName(k, type_op);
                    Handler.SaveOriginValues(_ws.Cells[row_type_op, column], tableName: ALL_OLD_VALUE);
                    _ws.Cells[row_type_op, column].Value2 = op_sec_type;
                    Handler.StoreEdit(_ws.Cells[row_type_op, column], tableName: MODIFICA);

                    // inserisco codice bilanciamento
                    string tmp_balance = siglaGradino.Substring(0, siglaGradino.Length - 1) + "CB";
                    int row_tmp_balance = _definedNames.GetRowByName(k, tmp_balance);
                    Handler.SaveOriginValues(_ws.Cells[row_tmp_balance, column], tableName: ALL_OLD_VALUE);
                    _ws.Cells[row_tmp_balance, column].Value2 = codBalance;
                    Handler.StoreEdit(_ws.Cells[row_tmp_balance, column], tableName: MODIFICA);
                }


                if (residuo > 0)
                {
                    int grad = int.Parse(siglaInfoTipo.Substring(siglaInfoTipo.Length - 5, 1));
                    if (grad < 4)
                    {
                        string tmp_str = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 2) + (grad + 1) + "TIPO";
                        int tmp_int = _definedNames.GetRowByName(lbSiglaEntitaPrinc.Tag, tmp_str);
                        Handler.SaveOriginValues(_ws.Cells[tmp_int, column], tableName: ALL_OLD_VALUE);
                        _ws.Cells[tmp_int, column].Value2 = _ws.Cells[row_tipo, column].Value2;
                        Handler.StoreEdit(_ws.Cells[tmp_int, column], tableName: MODIFICA);                        

                        tmp_str = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 2) + (grad + 1) + "E";
                        tmp_int = _definedNames.GetRowByName(lbSiglaEntitaPrinc.Tag, tmp_str);
                        Handler.SaveOriginValues(_ws.Cells[tmp_int, column], tableName: ALL_OLD_VALUE);
                        _ws.Cells[tmp_int, column].Value2 = residuo;
                        Handler.SaveOriginValues(_ws.Cells[tmp_int+1, column], tableName: ALL_OLD_VALUE);
                        _ws.Cells[tmp_int+1, column].Value2 = 0;
                        Handler.StoreEdit(_ws.Cells[tmp_int, column], tableName: MODIFICA);

                        /*
                        tmp_str = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 2) + (grad + 1) + "CB";
                        tmp_int = _definedNames.GetRowByName(lbSiglaEntitaPrinc.Tag, tmp_str);
                        Handler.SaveOriginValues(_ws.Cells[tmp_int, column], tableName: ALL_OLD_VALUE);
                        _ws.Cells[tmp_int, column].Value2 = codBalance;
                        Handler.StoreEdit(_ws.Cells[tmp_int, column], tableName: MODIFICA);
                        */
                    }
                }

                Sheet.Protected = true;

                traceBalance.Clear();
            }

            btnRipristina.Enabled = true;
        }

        private DataRow GetEntitaFromRange(Excel.Range rng)
        {
            DataRow o = null;
            if (_definedNames != null && _definedNames.IsDefined(rng.Row))
            {
                string nome = _definedNames.GetNameByAddress(rng.Row, rng.Column);
                string siglaEntita = nome.Split(Simboli.UNION[0])[0];

                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;

                if (categoriaEntita.Count > 0)
                    o = categoriaEntita[0].Row;
            }
            return o;
        }

        private string GetInfoFromRange(Excel.Range rng)
        {
            string o = null;
            if (_definedNames != null && _definedNames.IsDefined(rng.Row))
            {
                string nome = _definedNames.GetNameByAddress(rng.Row, rng.Column);
                //string siglaEntita = nome.Split(Simboli.UNION[0])[0];
                o = nome.Split(Simboli.UNION[0])[1];
            }
            return o;
        }

        private void FormBilanciamento_FormClosed(object sender, EventArgs e)
        {
            _ws.SelectionChange -= ChangeBalanceHours;
        }

        private void FormBilanciamento_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (Workbook.Repository[MODIFICA] != null && Workbook.Repository[MODIFICA].Rows.Count > 0)
            {
                DialogResult dr = MessageBox.Show("Salvare le modifiche apportate ai valori? Premere sì per inviare le modifiche al server, no per cancellarle definitivamente. Non sarà possibile recuperarle.", Simboli.NomeApplicazione + " - ATTENZIONE!!!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);

                if (dr == DialogResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }

                if (dr == DialogResult.Yes)
                    DataBase.SalvaModificheDB(MODIFICA);

                if (dr == DialogResult.No)
                    RipristinaValoriDaRepositori();
            }


            dgvBilanciamento.CellValidating -= ValidatePriority;
            dgvBilanciamento.CellValidated -= AfterValidation;

            Workbook.Repository.Remove(MODIFICA);
            Workbook.Repository.Remove(ALL_OLD_VALUE);
        }

        private void RipristinaValori_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Ripristinare i valori originali? Premere sì per continuare, no per lasciare i valori attuali.", Simboli.NomeApplicazione + " - ATTENZIONE!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                RipristinaValoriDaRepositori();
            }
        }

        private void RipristinaValoriDaRepositori()
        {
            DataTable dt = Workbook.Repository[tableName: ALL_OLD_VALUE];
            DataTable dtm = Workbook.Repository[tableName: MODIFICA];
            DataRow[] drs;

            Sheet.Protected = false;

            foreach (DataRow dr in dt.Rows)
            {
                string data = dr["Data"].ToString();
                string suffData = Date.GetSuffissoData(new DateTime(int.Parse(data.Substring(0, 4)), int.Parse(data.Substring(4, 2)), int.Parse(data.Substring(6, 2))));
                string suffOra = Date.GetSuffissoOra(int.Parse(data.Substring(8, 2)));
                int col = _definedNames.GetColFromDate(suffData, suffOra);
                int row = _definedNames.GetRowByName(dr["SiglaEntita"].ToString() + Simboli.UNION[0] + dr["SiglaInformazione"].ToString());

                _ws.Cells[row, col].Value2 = dr["Valore"];

                _ws.Cells[row, col].ClearComments();

                if (!dr["Commento"].Equals(""))
                    _ws.Cells[row, col].AddComment(dr["Commento"]);

                drs = dtm.Select("SiglaEntita = '" + dr["SiglaEntita"].ToString() + "' AND SiglaInformazione = '" +
                    dr["SiglaInformazione"].ToString() + "' AND Data = '" +
                    dr["Data"].ToString() + "'");
                if (drs.Count() > 0)
                {
                    dtm.Rows.Remove(drs.FirstOrDefault());
                }
            }

            Sheet.Protected = true;
            dt.Clear();

            btnRipristina.Enabled = false;
        }

        private void ValidatePriority(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.FormattedValue != "")
            {
                int value;
                int.TryParse(e.FormattedValue.ToString(), out value);
                if (e.ColumnIndex == 0 && value <= 0)
                {
                    MessageBox.Show("Inserire un numero intero maggiore di zero.");
                    e.Cancel = true;
                }
            }
        }

        private void AfterValidation(object sender, EventArgs e)
        {
            bool EnableModufy = false;
            for (int i = 0; i < dgvBilanciamento.RowCount; i++)
            {
                if (dgvBilanciamento[0, i].Value.ToString() != "")
                {
                    EnableModufy = true;
                }
            }

            btnModifica.Enabled = EnableModufy;
        }

        private void FormBilanciamento_Load(object sender, EventArgs e)
        {
            dgvBilanciamento.CellValidating += ValidatePriority;
            dgvBilanciamento.CellValidated += AfterValidation;
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            Close();
        }

    }
}