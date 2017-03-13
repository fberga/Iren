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
        public const string MODIFICA = "FormBilanciamentoModifica";
        public const string ALL_OLD_VALUE = "FormBilanciamentoRipristino";

        //private DefinedNames _definedNames;
        private Excel.Worksheet _ws;

        private Dictionary<string, DefinedNames> _categoriaNomiDefiniti = new Dictionary<string, DefinedNames>();
        // SiglaEntità NomeFoglio
        Dictionary<string, string> EntitySheet;
        
        private string _categoriaEntitaPrincipale = "";
        private string _siglaEntitaPrincipale = "";

        private Excel.Range _origRng;

        private DataTable _dtEntitaGradinoDisp;

        // Mi permette di tenere traccia dell'operazione che devo apllicare
        // Se VEN filtro per favoli positivi, se ACQ per negetivi
        string op_sec_type;

        public FormBilanciamento(Excel.Worksheet ws, Excel.Range rng)
        {
            InitializeComponent();
            this.Text = Simboli.NomeApplicazione + " - Bilanciamento";

            foreach (Excel.Worksheet s in Workbook.CategorySheets)
            {
                string siglaCategoria = Workbook.Repository[DataBase.TAB.CATEGORIA].AsEnumerable()
                    .Where(r => r["DesCategoria"].Equals(s.Name))
                    .Select(r => r["SiglaCategoria"].ToString())
                    .FirstOrDefault();

                _categoriaNomiDefiniti.Add(siglaCategoria, new DefinedNames(s.Name, DefinedNames.InitType.All));
            }

            _ws = ws;

            _ws.SelectionChange += ChangeBalanceHours;
            //_definedNames = new DefinedNames(_ws.Name, DefinedNames.InitType.All);


            // Previene errori in caso di chiusura form non previsti -> errori non gestiti
            if (Workbook.Repository[MODIFICA] == null)
                Workbook.Repository.Add(Workbook.Repository.CreaTabellaModifica(MODIFICA));
            else
                Workbook.Repository[MODIFICA].Reset();

            if (Workbook.Repository[ALL_OLD_VALUE] == null)
                Workbook.Repository.Add(Workbook.Repository.CreaTabellaRipristinaIncremento(ALL_OLD_VALUE));
            else
                Workbook.Repository[ALL_OLD_VALUE].Reset();

            

            //if (!CheckPrecondition())
            //{
            //    //MessageBox.Show("Non sono state selezionate celle valide. Selezionare celle offerta energia.");
            //    return;
            //}

            
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

            categoriaEntita.RowFilter = "SiglaEntita <> '" + _siglaEntitaPrincipale + "' AND Zona = '" + lbZonaUPprinc.Text + "' AND IdApplicazione = " + Workbook.IdApplicazione + " AND Gerarchia is null";

            //dgvBilanciamento.DataSource = categoriaEntita;

            int hCount = Target.Columns.Count;

            DataTable dt = new DataTable()
            {
                Columns =
                {
                    {"Priorità", typeof(string)},
                    {"VaiA", typeof(object)},
                    {"DesEntita", typeof(string)},
                    {"SiglaEntita", typeof(string)},
                    {"SiglaCategoria", typeof(string)}
                    //{"SiglaInformazione", typeof(string)}
                }
            };
            
            foreach (Excel.Range h in Target)
            {
                dt.Columns.Add(new DataColumn(""+h.Column, typeof(double)));
                _dtEntitaGradinoDisp.Columns.Add(new DataColumn("" + h.Column, typeof(string)));
            }
            

            string siglaInfoEnergia = GetInfoFromRange(Target);
            string siglaInfoTipo = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 1) + "TIPO";

            //int row = _categoriaNomiDefiniti[_categoriaEntitaPrincipale].GetRowByName(_siglaEntitaPrincipale, siglaInfoTipo);

            //cerco acq - ven
            //int priorityvalue = categoriaEntita.Count;


            // Imposto il valore una sola volta
            int row_tipo = _categoriaNomiDefiniti[_categoriaEntitaPrincipale].GetRowByName(_siglaEntitaPrincipale, siglaInfoTipo);
            op_sec_type = "";
            if (_ws.Cells[row_tipo, Target.Column].Value2.Equals("VEN"))
            {
                op_sec_type = "ACQ";
            }
            else
                op_sec_type = "VEN";

            foreach (DataRowView entita in categoriaEntita)
            {

                int row = _categoriaNomiDefiniti[entita["SiglaCategoria"].ToString()].GetRowByName(entita["SiglaEntita"], siglaInfoTipo);

                DataRow offSecondaria = dt.NewRow();
                
                DataRow gradiniDisp = _dtEntitaGradinoDisp.NewRow();

                gradiniDisp["SiglaEntita"] = entita["SiglaEntita"];

                //offSecondaria["Priorità"] = priorityvalue;
                offSecondaria["SiglaEntita"] = entita["SiglaEntita"];
                offSecondaria["DesEntita"] = entita["DesEntita"];
                offSecondaria["SiglaCategoria"] = entita["SiglaCategoria"];

                foreach (Excel.Range h in Target)
                {
                    Excel.Worksheet ws_tmp = Workbook.Application.Worksheets[_categoriaNomiDefiniti[offSecondaria["SiglaCategoria"].ToString()].Sheet];
                    /*
                    string siglaInfo = ws_tmp.Cells[row, h.Column].Value.Equals("VEN") ? "TEMP_OMI11" : "TEMP_OMI10";

                    int row1 = _categoriaNomiDefiniti[entita["SiglaCategoria"].ToString()].GetRowByName(entita["SiglaEntita"], siglaInfo);

                    offSecondaria["" + h.Column] = ws_tmp.Cells[row1, h.Column].Value2;
                    */

                    for (int j = 1; j < 5; j++)
                    {
                        string siglaGradino = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 2) + j + "E";
                        int row2 = _categoriaNomiDefiniti[entita["SiglaCategoria"].ToString()].GetRowByName(entita["SiglaEntita"], siglaGradino);
                        if (ws_tmp.Cells[row2, h.Column].Value2 == null)
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

            //DataGridViewColumn combo = new DataGridViewColumn();
            //combo.HeaderText = "ppp";
            //combo.Name = "combo";

            dgvBilanciamento.Columns["DesEntita"].ReadOnly = true;
            dgvBilanciamento.Columns["SiglaEntita"].Visible = false;
            dgvBilanciamento.Columns["SiglaCategoria"].Visible = false;
            dgvBilanciamento.Columns["VaiA"].Width = 150;
            dgvBilanciamento.Columns["DesEntita"].Width = 120;
            foreach (Excel.Range h in Target)
            {
                dgvBilanciamento.Columns["" + h.Column].HeaderText = Date.GetSuffissoOra(h.Column - _categoriaNomiDefiniti[_categoriaEntitaPrincipale].GetFirstCol());
                dgvBilanciamento.Columns["" + h.Column].ReadOnly = true;
            }

            _origRng = Target;

            FormBilanciamento_Shown(null, null);
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
            DataRow entita = GetEntitaFromRange(Target);
            if (entita == null)
            {
                labelError.Text = "ERRORE: Posizionarsi sul foglio della selezione.";
                return false;
            }
            _categoriaEntitaPrincipale = entita["SiglaCategoria"].ToString();

            //imposto valori da visualizzare
            lbSiglaEntitaPrinc.Text = entita["DesEntita"].ToString();
            _siglaEntitaPrincipale = entita["SiglaEntita"].ToString();

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

            //precondizione: Selezionata riga energia
            string siglaInformazione = GetInfoFromRange(Workbook.Application.Selection);
            if (!Regex.IsMatch(siglaInformazione, @"OFFERTA_MI\d_G\dE"))
            {
                labelError.Text = "ATTENZIONE: Si possono selezionare solo righe relative alla quantità di energia.";
                return false;
            }

            string tipo_op = siglaInformazione.Substring(0, siglaInformazione.Length - 1) + "TIPO";
            int row_tipo_op = _categoriaNomiDefiniti[_categoriaEntitaPrincipale].GetRowByName(_siglaEntitaPrincipale, tipo_op);
            
            foreach (Excel.Range row in Target.Rows)
            {
                if (!_categoriaNomiDefiniti[_categoriaEntitaPrincipale].IsEditable(row.Row))
                {
                    labelError.Text = "ERRORE: Il range selezionato contiene delle righe non modificabili.";
                    return false;
                }
            }

            string compare_op = "";
            // Controllo che tutte le ora selezionate contengano lo stesso tipo di operazione ACQ o VEN
            foreach (Excel.Range row in Target.Columns)
            {
                if (string.IsNullOrEmpty(compare_op) )
                {
                    compare_op = _ws.Cells[row_tipo_op, row.Column].Value2;
                }
                else
                {
                    if (!compare_op.Equals(_ws.Cells[row_tipo_op, row.Column].Value2))
                    {
                        labelError.Text = "ERRORE: Si può effettuare un'operazione di bilanciamento su più ore solo se l'operazione (ACQ/VEN) è la stessa per tutte le ore della selezione.";
                        return false;
                    }
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

            //imposto valori da visualizzare
            lbSiglaEntitaPrinc.Text = entita["DesEntita"].ToString();
            _siglaEntitaPrincipale = entita["SiglaEntita"].ToString();

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
            EntitySheet = new Dictionary<string, string>();
            DataView dv = ((DataTable)dgvBilanciamento.DataSource).AsDataView();
            Dictionary<Tuple<string, string>, double> traceBalance = new Dictionary<Tuple<string, string>, double>();

            dv.RowFilter = "Priorità is not null AND VaiA is null";
            if (dv.Count > 0)
            {
                labelError.Text = "ERRORE: Tutte le righe a cui è stata assegnata una priorità devono avere un campbo di calcolo 'VaiA' selezionato e viceversa.";
                return;
            }

            dv.RowFilter = " Priorità is not null AND VaiA is not null";
            dv.Sort = "Priorità ASC";

            if (dv.Count == 0)
            {
                labelError.Text = "ERRORE: Attribuire almeno una priorità ed un valore 'VaiA' di riferimento.";
                return;
            }

            int column;

            string siglaInfoEnergia = GetInfoFromRange(_origRng);
            
            string siglaInfoTipo = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 1) + "TIPO";
            string siglaPrezzo = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 1) + "P";
            int row_tipo = _categoriaNomiDefiniti[_categoriaEntitaPrincipale].GetRowByName(_siglaEntitaPrincipale, siglaInfoTipo);
            int row_prezzo = _categoriaNomiDefiniti[_categoriaEntitaPrincipale].GetRowByName(_siglaEntitaPrincipale, siglaPrezzo);

            int time = (int)DateTime.Now.TimeOfDay.TotalSeconds;
            
            EntitySheet.Add(_categoriaEntitaPrincipale, _categoriaNomiDefiniti[_categoriaEntitaPrincipale].Sheet);

            foreach (Excel.Range cell in _origRng )
            {
                column = cell.Column;

                double residuo = _ws.Cells[cell.Row, column].Value2;

                // Calcolo i contributi
                foreach (DataRowView dtv in dv)
                {
                    double secondaryOffert;
                    bool res_parse = double.TryParse(dtv["" + column].ToString(), out secondaryOffert);
                    
                    if (!res_parse)
                        continue;
                    else if (secondaryOffert == 0)
                        continue;

                    if (residuo >= secondaryOffert)
                    {
                        residuo -= secondaryOffert;

                        // salvo gli UP coinvolti nel bilanciamento
                        traceBalance.Add(Tuple.Create(dtv["SiglaCategoria"].ToString(), dtv["SiglaEntita"].ToString()), secondaryOffert);

                        if (residuo == 0)
                            break;
                    }
                    else
                    {
                        traceBalance.Add(Tuple.Create(dtv["SiglaCategoria"].ToString(), dtv["SiglaEntita"].ToString()), residuo);
                        residuo = 0;
                        break;
                    }
                }

                if (residuo == _ws.Cells[cell.Row, column].Value2)
                {
                    labelError.Text = "Con i dati inseriti/selezionati il bilanciamento non ha sortito nessun effetto.";
                    return;
                }
                
                // Aggiorno il foglio con i contributi per l'ora esaminata
                Sheet.Protected = false;
                int column_con = column - _categoriaNomiDefiniti[_categoriaEntitaPrincipale].GetFirstCol();
                string codBalance = (column_con < 10 ? "0" + column_con : "" + column_con) + time;

                //scrivo il codice bilanciamento nell'UP principale
                string sigla_balance = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 1) + "CB";
                int row_balance = _categoriaNomiDefiniti[_categoriaEntitaPrincipale].GetRowByName(_siglaEntitaPrincipale, sigla_balance);

                Handler.SaveOriginValues(_ws.Cells[row_prezzo, column], ALL_OLD_VALUE, _categoriaEntitaPrincipale);
                _ws.Cells[row_prezzo, column].Value2 = "";
                Handler.StoreEdit(_ws.Cells[row_prezzo, column], tableName: MODIFICA);

                Handler.SaveOriginValues(_ws.Cells[row_balance, column], ALL_OLD_VALUE, _categoriaEntitaPrincipale);
                _ws.Cells[row_balance, column].Value2 = codBalance;
                Handler.StoreEdit(_ws.Cells[row_balance, column], tableName: MODIFICA);

                // ciclo per tutti gli Up coinvolti nel bilanciamento
                foreach (Tuple<string, string> k in traceBalance.Keys)
                {
                    Excel.Worksheet ws_tmp = Workbook.Application.Worksheets[_categoriaNomiDefiniti[k.Item1].Sheet];

                    string siglaGradino = _dtEntitaGradinoDisp.Rows.Find(k.Item2).ItemArray[1].ToString();
                    int tmp_row = _categoriaNomiDefiniti[k.Item1].GetRowByName(k.Item2, siglaGradino);

                    if (!EntitySheet.ContainsKey(k.Item1))
                        EntitySheet.Add(k.Item1, ws_tmp.Name);

                    Handler.SaveOriginValues(ws_tmp.Cells[tmp_row, column], ALL_OLD_VALUE, k.Item1);
                    ws_tmp.Cells[tmp_row, column].Value2 = traceBalance[k];
                    Handler.StoreEdit(ws_tmp.Cells[tmp_row, column], tableName: MODIFICA);
                    // modifico il prezzo e lo imposto = 0
                    Handler.SaveOriginValues(ws_tmp.Cells[tmp_row + 1, column], ALL_OLD_VALUE, k.Item1);
                    ws_tmp.Cells[tmp_row + 1, column].Value2 = ""; // campbiare secondo le proprietà dei nomi cella
                    Handler.StoreEdit(ws_tmp.Cells[tmp_row + 1, column], tableName: MODIFICA);

                    // modifico la label ACQ/VEN
                    string type_op = siglaGradino.Substring(0, siglaGradino.Length - 1) + "TIPO";
                    int row_type_op = _categoriaNomiDefiniti[k.Item1].GetRowByName(k.Item2, type_op);
                    Handler.SaveOriginValues(ws_tmp.Cells[row_type_op, column], ALL_OLD_VALUE, k.Item1);
                    ws_tmp.Cells[row_type_op, column].Value2 = op_sec_type;
                    Handler.StoreEdit(ws_tmp.Cells[row_type_op, column], tableName: MODIFICA);

                    // inserisco codice bilanciamento
                    string tmp_balance = siglaGradino.Substring(0, siglaGradino.Length - 1) + "CB";
                    int row_tmp_balance = _categoriaNomiDefiniti[k.Item1].GetRowByName(k.Item2, tmp_balance);
                    Handler.SaveOriginValues(ws_tmp.Cells[row_tmp_balance, column], ALL_OLD_VALUE, k.Item1);
                    ws_tmp.Cells[row_tmp_balance, column].Value2 = codBalance;
                    Handler.StoreEdit(ws_tmp.Cells[row_tmp_balance, column], tableName: MODIFICA);
                }


                if (residuo > 0)
                {
                    int grad = int.Parse(siglaInfoTipo.Substring(siglaInfoTipo.Length - 5, 1));
                    if (grad < 4)
                    {
                        // Modifico l'offerta iniziale secondo la quota bilanciata - Salvo le info necessarie per il salva e il ripristina
                        Handler.SaveOriginValues(_ws.Cells[cell.Row, column], ALL_OLD_VALUE, _categoriaEntitaPrincipale);
                        _ws.Cells[cell.Row, column].Value2 -= residuo;
                        Handler.StoreEdit(_ws.Cells[cell.Row, column], tableName: MODIFICA);


                        // Inserisco nelgradino successivo un'operazione per il residuo
                        string tmp_str = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 2) + (grad + 1) + "TIPO";
                        int tmp_int = _categoriaNomiDefiniti[_categoriaEntitaPrincipale].GetRowByName(_siglaEntitaPrincipale, tmp_str);
                        Handler.SaveOriginValues(_ws.Cells[tmp_int, column], ALL_OLD_VALUE, _categoriaEntitaPrincipale);
                        _ws.Cells[tmp_int, column].Value2 = _ws.Cells[row_tipo, column].Value2;
                        Handler.StoreEdit(_ws.Cells[tmp_int, column], tableName: MODIFICA);                        

                        tmp_str = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 2) + (grad + 1) + "E";

                        tmp_int = _categoriaNomiDefiniti[_categoriaEntitaPrincipale].GetRowByName(_siglaEntitaPrincipale, tmp_str);
                        Handler.SaveOriginValues(_ws.Cells[tmp_int, column], ALL_OLD_VALUE, _categoriaEntitaPrincipale);
                        _ws.Cells[tmp_int, column].Value2 = residuo;
                        Handler.StoreEdit(_ws.Cells[tmp_int, column], tableName: MODIFICA);

                        Handler.SaveOriginValues(_ws.Cells[tmp_int + 1, column], ALL_OLD_VALUE, _categoriaEntitaPrincipale);
                        _ws.Cells[tmp_int + 1, column].Value2 = "";
                        Handler.StoreEdit(_ws.Cells[tmp_int + 1, column], tableName: MODIFICA);

                        /*
                        tmp_str = siglaInfoEnergia.Substring(0, siglaInfoEnergia.Length - 2) + (grad + 1) + "CB";
                        tmp_int = _definedNames.GetRowByName(_siglaEntitaPrincipale, tmp_str);
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
            DefinedNames d = _categoriaNomiDefiniti
                .Where(kv => kv.Value.Sheet == rng.Worksheet.Name)
                .Select(kv => kv.Value)
                .FirstOrDefault();
            
            

            DataRow o = null;
            if (d != null && d.IsDefined(rng.Row))
            {
                string nome = d.GetNameByAddress(rng.Row, rng.Column);
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
            DefinedNames d = _categoriaNomiDefiniti
                .Where(kv => kv.Value.Sheet == rng.Worksheet.Name)
                .Select(kv => kv.Value)
                .FirstOrDefault();

            string o = null;
            if (d != null && d.IsDefined(rng.Row))
            {
                string nome = d.GetNameByAddress(rng.Row, rng.Column);
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
            /*
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
            */

            if (Workbook.Repository[MODIFICA] != null && Workbook.Repository[MODIFICA].Rows.Count > 0)
            {
                DataBase.SalvaModificheDB(MODIFICA);
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
            // TODO: DA testare
            
            foreach (DataRow dr in dt.Rows)
            {
                string data = dr["Data"].ToString();
                string suffData = Date.GetSuffissoData(new DateTime(int.Parse(data.Substring(0, 4)), int.Parse(data.Substring(4, 2)), int.Parse(data.Substring(6, 2))));
                string suffOra = Date.GetSuffissoOra(int.Parse(data.Substring(8, 2)));

                //traceBalance.Keys.
                int col = _categoriaNomiDefiniti[dr["SiglaCategoria"].ToString()].GetColFromDate(suffData, suffOra);
                int row = _categoriaNomiDefiniti[dr["SiglaCategoria"].ToString()].GetRowByName(dr["SiglaEntita"].ToString() + Simboli.UNION[0] + dr["SiglaInformazione"].ToString());

                Excel.Worksheet ws_tmp = Workbook.Application.Worksheets[EntitySheet[dr["SiglaCategoria"].ToString()]];

                ws_tmp.Cells[row, col].Value2 = dr["Valore"].ToString() == "0" ? "" : dr["Valore"];
                ws_tmp.Cells[row, col].ClearComments();

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
            if (e.ColumnIndex == dgvBilanciamento.Columns["Priorità"].Index && e.FormattedValue != "")
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
            // TODO CHANGE IF
            if (sender == null)
            {
                bool EnableModufy = false;
                for (int i = 0; i < dgvBilanciamento.RowCount; i++)
                {
                    if (dgvBilanciamento[0, i].Value.ToString() != "")
                    {
                        EnableModufy = true;
                    }
                }

                
            }
            // rimettere dentro l'if
            btnModifica.Enabled = true;
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

        private void FormBilanciamento_Shown(object sender, EventArgs e)
        {
            if (dgvBilanciamento.DataSource != null)
            {
                for (int i = 0; i < dgvBilanciamento.Rows.Count; i++)
                {
                    DataGridViewComboBoxCell cellCombo = new DataGridViewComboBoxCell();

                    string siglaEntita = dgvBilanciamento["SiglaEntita", i].Value.ToString();
                    string siglaInfo = _dtEntitaGradinoDisp.AsEnumerable()
                        .Where(r => r["SiglaEntita"].Equals(siglaEntita))
                        .Select(r => r[1].ToString())
                        .FirstOrDefault();

                    DefinedNames df = _categoriaNomiDefiniti[dgvBilanciamento["SiglaCategoria", i].Value.ToString()];

                    Dictionary<string, int> gt = OfferteMIHelper.GetGOTODictionary(siglaEntita, siglaInfo, df);

                    if (gt != null)
                    {
                        cellCombo.DataSource = new BindingSource(gt, null);
                        cellCombo.ValueMember = "Value";
                        cellCombo.DisplayMember = "Key";
                        dgvBilanciamento[1, i] = cellCombo;
                    }

                    // seleziono il primo valore della combo
                    //dgvBilanciamento[1, i].Selected = true;
                    //dgvBilanciamento[1, i].Value = gt.First().Key;
                }
            }
        }

        private void dgvBilanciamento_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dgvBilanciamento.Columns["VaiA"].Index)
            {
                //il foglio è unico
                int col = dgvBilanciamento.Columns["SiglaCategoria"].Index;
                int row_calc = (int)dgvBilanciamento.CurrentCell.Value;
                Excel.Worksheet ws_tmp = Workbook.Application.Worksheets[_categoriaNomiDefiniti[dgvBilanciamento[col,e.RowIndex].Value.ToString()].Sheet];
                foreach (DataGridViewColumn cc in dgvBilanciamento.Columns)
                {  
                    if (Regex.IsMatch(cc.HeaderText, @"H\d+"))
                    {
                        int col_calc;
                        int.TryParse(cc.Name, out col_calc);
                        double cell_value;
                        double value_DataGrid = 0;
                        if (string.IsNullOrEmpty(ws_tmp.Cells[row_calc, col_calc].Value2.ToString()) )
                            cell_value = 0;
                        else
                        {
                            bool result = double.TryParse( ws_tmp.Cells[row_calc, col_calc].Value2.ToString(), out cell_value );
                            if (!result)
                                cell_value = 0;
                            else
                            {
                                if (op_sec_type.Equals("VEN"))
                                {
                                    if (cell_value < 0)
                                        value_DataGrid = 0;
                                    else
                                        value_DataGrid = cell_value;
                                }
                                else if (op_sec_type.Equals("ACQ"))
                                {
                                    if (cell_value > 0)
                                        value_DataGrid = 0;
                                    else
                                        value_DataGrid = Math.Abs(cell_value);
                                }
                                else
                                    value_DataGrid = 0;
                            }

                        }
                        dgvBilanciamento[cc.Index, e.RowIndex].Value = value_DataGrid;
                    }
                }
            }
        }

        private void dgvBilanciamento_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvBilanciamento.CurrentCell is DataGridViewComboBoxCell)
            {
                dgvBilanciamento.CommitEdit(DataGridViewDataErrorContexts.Commit);
                dgvBilanciamento.EndEdit();
            }
        }

        // evento utile a far visualizzare subito le opzioni delle tendina al primo click
        private void dgvBilanciamento_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            // filtro solo per la colonna delle comboBox
            if (e.ColumnIndex == dgvBilanciamento.Columns["VaiA"].Index)
            {
                dgvBilanciamento.BeginEdit(false);
                if (this.dgvBilanciamento.EditingControl != null
                    && this.dgvBilanciamento.EditingControl is ComboBox)
                {
                    ComboBox cmb = this.dgvBilanciamento.EditingControl as ComboBox;
                    cmb.DroppedDown = true;
                }
            }
        }


    }
}