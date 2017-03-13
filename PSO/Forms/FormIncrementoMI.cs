using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Forms
{
    public partial class FormIncrementoMI : Form
    {
        public const string MODIFICA = "FormIncrementoMiModifica";
        public const string ALL_OLD_VALUE = "FormIncrementoMiRipristino";
        

        #region Variabili
        //private object[,] _origVal;
        //List<bool> _comments;
        List<string> _labelVenAcq = new List<string>();

        private Excel.Range _origRng;

        private DefinedNames _definedNames;

        private Excel.Worksheet _ws;

        private double? _percentage;
        private double? _increment;

        private bool _selectionIsCorrect = false;
        private bool _valuesAreCorrect = false;

        // Dizionario item combobox / riga
        //private Dictionary<string, int> offerta_dictionary = new Dictionary<string, int>();
        // Dizionario riga combo /riga calcolo
        //private Dictionary<string, string> calcolo_dictionary = new Dictionary<string, string>();

        private Dictionary<string, int> _gotoDictionary = new Dictionary<string, int>();

        private bool is_price = false;
        private bool is_quantity = false;

        #endregion

        #region Costruttore

        public FormIncrementoMI(Excel.Worksheet ws, Excel.Range rng)
        {
            InitializeComponent();
            this.Text = Simboli.NomeApplicazione + " - Incremento";

            _ws = ws;

            _ws.SelectionChange += ChangeSelectionToIncrement;
            _definedNames = new DefinedNames(_ws.Name, DefinedNames.InitType.All);


            // Previene errori in caso di chiusura form non previsti -> errori non gestiti
            if (Workbook.Repository[MODIFICA] == null)
                Workbook.Repository.Add(Workbook.Repository.CreaTabellaModifica(MODIFICA));
            else
                Workbook.Repository[MODIFICA].Reset();

            if (Workbook.Repository[ALL_OLD_VALUE] == null)
                Workbook.Repository.Add(Workbook.Repository.CreaTabellaRipristinaIncremento(ALL_OLD_VALUE));
            else
                Workbook.Repository[ALL_OLD_VALUE].Reset();
            
            btnRipristina.Enabled = false;
            btnApplica.Enabled = false;
            ChangeSelectionToIncrement(Workbook.Application.Selection);
        }

        #endregion

        #region Eventi

        private void ChangeSelectionToIncrement(Excel.Range Target)
        {
            // Descrizione Combo / Riga Calcolo
            //offerta_dictionary = new Dictionary<string, int>();
            // SiglaInformazione riga selezionata / SiglaInformazione riga calcolo
            //calcolo_dictionary = new Dictionary<string, string>();

            _gotoDictionary = new Dictionary<string, int>();

            btnApplica.Enabled = false;

            comboBox_VaiA.DataSource = null;
            comboBox_applicaA.DataSource = null;

            groupQuantità.Visible = false;
            groupPrezzo.Visible = false;

            lbErrore.Text = "";
            lbErrore.ForeColor = Color.Red;
            btnApplica.Enabled = false;
            _selectionIsCorrect = false;

            if (Target.Rows.Count > 1)
            {
                foreach (Excel.Range row in Target.Rows)
                {
                    if (row.EntireRow.Hidden)
                    {
                        lbErrore.Text = "ERRORE: Nel range selezionato ci sono righe nascoste.";
                        return;
                    }
                }
                lbErrore.ForeColor = Color.DarkOrange;
                lbErrore.Text = "ATTENZIONE: Nel range selezionato ci sono più righe.";
            }
            foreach (Excel.Range row in Target.Rows)
            {
                if (!_definedNames.IsEditable(row.Row))
                {
                    lbErrore.Text = "ERRORE: Il range selezionato contiene delle righe non modificabili.";
                    return;
                }
            }

            int firstCol = _definedNames.GetFirstCol();
            if (Target.Column < firstCol + Simboli.GetMarketOffsetMI(Workbook.Mercato, Workbook.DataAttiva))
            {
                lbErrore.Text = "ERRORE: Il range selezionato contiene celle appartenenti a mercati chiusi.";
                return;
            }

            /* Controllo se la riga selezionata è corretta per l'operazione richiesta */
            string name_col_row_selected = _definedNames.GetNameByRow(Target.Row).FirstOrDefault();
            string siglaEntita = name_col_row_selected.Split('.').First();
            string siglaInformazione = name_col_row_selected.Split('.').Last();
            is_price = is_Name_Match_Price(name_col_row_selected);
            is_quantity = is_Name_Match_Quantity(name_col_row_selected);

            //DataView definizioneOfferta = Workbook.Repository[DataBase.TAB.DEFINIZIONE_OFFERTA].DefaultView;

            //definizioneOfferta.RowFilter = "SiglaEntita ='" + siglaEntita + "' AND SiglaInformazione = '" + siglaInformazione + "' AND IdMercato = " + Workbook.Mercato.Substring(2, Workbook.Mercato.Length - 2);

            //if (definizioneOfferta.Count == 0)
            //{
            //    lbErrore.Text = "ERRORE: Non ci sono opzioni attive per questa funzione.";
            //    return;
            //}

            //DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            //DataTable entitaInformazione = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE];

            //foreach (DataRowView offerta in definizioneOfferta)
            //{
            //    string desInformazioneCombo = entitaInformazione.AsEnumerable()
            //        .Where(r => r["SiglaEntita"].Equals(offerta["SiglaEntita"])
            //                 && (r["SiglaEntitaRif"] is DBNull || r["SiglaEntitaRif"].Equals(offerta["SiglaEntitaCombo"]))
            //                 && r["SiglaInformazione"].Equals(offerta["SiglaInformazioneCombo"]))
            //        .Select(r => r["DesInformazione"].ToString())
            //        .FirstOrDefault();

            //    object entitaCalcolo = offerta["SiglaEntitaCalcolo"] is DBNull ? offerta["SiglaEntitaCombo"] : offerta["SiglaEntitaCalcolo"];
            //    object infoCalcolo = offerta["SiglaInformazioneCalcolo"] is DBNull ? offerta["SiglaInformazioneCombo"] : offerta["SiglaInformazioneCalcolo"];

            //    _gotoDictionary.Add(desInformazioneCombo, _definedNames.GetRowByName(entitaCalcolo, infoCalcolo));
            //}

            _gotoDictionary = OfferteMIHelper.GetGOTODictionary(siglaEntita, siglaInformazione, _definedNames);

            if (_gotoDictionary == null)
            {
                lbErrore.Text = "ERRORE: Non ci sono opzioni attive per questa funzione.";
                return;
            }

            if (is_price)
            {
                groupQuantità.Visible = false;
                groupPrezzo.Visible = true;

                comboBox_applicaA.DataSource = new BindingSource(_gotoDictionary, null);
                comboBox_applicaA.ValueMember = "Value";
                comboBox_applicaA.DisplayMember = "Key";
                comboBox_applicaA.SelectedIndex = -1;
            }
            else if (is_quantity)
            {
                groupQuantità.Visible = true;
                groupPrezzo.Visible = false;

                comboBox_VaiA.DataSource = new BindingSource(_gotoDictionary, null);
                comboBox_VaiA.ValueMember = "Value";
                comboBox_VaiA.DisplayMember = "Key";
                comboBox_VaiA.SelectedIndex = -1;
            }
            else
            {
                lbErrore.Text = "ERRORE: Il range selezionato non si riferisce a quantità o prezzi.";
                return;
            }

            _selectionIsCorrect = true;

            if (_valuesAreCorrect)
                btnApplica.Enabled = true;

            btnRipristina.Enabled = false;

            _origRng = Target;   

            Handler.SaveOriginValues(Target, tableName: ALL_OLD_VALUE);
            if (is_price)
            {
                Excel.Range rng = _ws.Range[_ws.Cells[Target.Row - 2, Target.Column], _ws.Cells[Target.Row - 2, Target.Column + Target.Columns.Count - 1]];
                Handler.SaveOriginValues(rng, tableName: ALL_OLD_VALUE);
            }
            if (is_quantity)
            {
                Excel.Range rng = _ws.Range[_ws.Cells[Target.Row - 1, Target.Column], _ws.Cells[Target.Row - 1, Target.Column + Target.Columns.Count - 1]];
                Handler.SaveOriginValues(rng, tableName: ALL_OLD_VALUE);
            }

        }

        private void TipoIncermento_checkedChanged(object sender, EventArgs e)
        {
            RadioButton rb = sender as RadioButton;

            if (rb.Name == "rdbPercentuale")
            {
                txtPercentuale.Enabled = true;
                txtValore.Enabled = false;
            }
            else
            {
                txtPercentuale.Enabled = false;
                txtValore.Enabled = true;
            }
        }

        private void TextElements_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;

            if (txt.Text == "")
            {
                /***************************************************************************************************/
                // _valuesAreCorrect = false;
                _valuesAreCorrect = true;
                /***************************************************************************************************/
                StateChanged_enableButton();
                return;
            }
           
            double val;
            string text = txt.Text.Replace(".", ",");


            //if (!Double.TryParse(text, out val))
            if (!Double.TryParse(text, out val) || val<0)
            {
                _valuesAreCorrect = false;
                StateChanged_enableButton();
                return;
            }

            if (txt.Name == "txtPercentuale")
            {
                _percentage = val;
                _increment = null;
            }
            else if (txt.Name == "txtValore")
            {
                _increment = val;
                _percentage = null;
            }

            _valuesAreCorrect = true;

            StateChanged_enableButton();
        }

        private void TextElements_EnabledChanged(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;
            if (txt.Enabled == true)
                TextElements_TextChanged(sender, e);
        }

        private void btnApplica_Click(object sender, EventArgs e)
        {
            lbErrore.Text = "";

            Sheet.Protected = false;

            foreach (Excel.Range rng in _origRng)
            {
                double result_tmp = 0;
                if (is_price)
                {
                    if (_ws.Cells[comboBox_applicaA.SelectedValue, rng.Column].Value2 != null)
                        result_tmp = _ws.Cells[comboBox_applicaA.SelectedValue, rng.Column].Value2;

                    if (rdbPercentuale.Checked)
                    {
                        if (string.IsNullOrEmpty(txtPercentuale.Text))
                            _percentage = 0.0;

                        result_tmp = result_tmp + (result_tmp * (_percentage.Value / 100));
                    }
                    else if (rdbIncremento.Checked)
                    {
                        if (string.IsNullOrEmpty(txtValore.Text))
                            _increment = 0.0;

                        result_tmp = result_tmp + _increment.Value;
                    }
                    if (result_tmp < 0)
                    {
                        lbErrore.Text = "ERRORE: Con i valori inseriti si ottiene un prezzo negativo. Verificare i dati inseriti.";
                    return;
                    }
                    else
                    {
                        if (result_tmp > 0)
                            rng.Value2 = result_tmp;
                    }
                        
                    // Modifica 21/02/2017 INIZIO Il campo prezzo può avere valori ACQ/VEN selezionati dall'utente e non modificabili
                    /*
                    if (result_tmp > 0)
                    {
                        _ws.Cells[rng.Row - 2, rng.Column].Value = "VEN";
                    }
                    else if (result_tmp < 0)
                    {
                        _ws.Cells[rng.Row - 2, rng.Column].Value = "ACQ";
                    }
                    else
                    {
                        _ws.Cells[rng.Row - 2, rng.Column].Value = "";
                    }
                    */
                    // Modifica 21/02/2017 FINE 
                }
                else if (is_quantity)
                {
                    rng.Value2 = "";
                    _ws.Cells[comboBox_VaiA.SelectedValue, rng.Column].Calculate();
                    result_tmp = _ws.Cells[comboBox_VaiA.SelectedValue, rng.Column].Value2;
                    if( result_tmp != 0 )
                        rng.Value2 = Math.Abs(result_tmp);

                    //_ws.Cells[offerta_dictionary[comboBox_VaiA.Text.ToString()], rng.Column].Calculate();
                    //result_tmp = _ws.Cells[offerta_dictionary[comboBox_VaiA.Text.ToString()], rng.Column].Value2;
                    //if (_ws.Cells[offerta_dictionary[comboBox_VaiA.Text.ToString()], rng.Column].Value2 != null)
                    //    rng.Value = Math.Abs(result_tmp);
                    //else
                    //  rng.Value = 0;

                    if (result_tmp < 0)
                    {
                        _ws.Cells[rng.Row - 1, rng.Column].Value2 = "ACQ";
                    }
                    else // valore di default
                    {
                        _ws.Cells[rng.Row - 1, rng.Column].Value2 = "VEN";
                    }
                        
                }
            }           

            Handler.StoreEdit(_origRng, tableName: MODIFICA);

            var a = Workbook.Repository;

            _origRng.Select();
            btnRipristina.Enabled = true;

            Sheet.Protected = true;
        }

        private void FormIncremento_FormClosed(object sender, FormClosedEventArgs e)
        {
            _ws.SelectionChange -= ChangeSelectionToIncrement;
        }

        private void RipristinaValori_Click(object sender, EventArgs e)
        {
            lbErrore.Text = "";

            if (MessageBox.Show("Ripristinare i valori originali? Premere sì per continuare, no per lasciare i valori attuali.", Simboli.NomeApplicazione + " - ATTENZIONE!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {

                RipristinaValoriDaRepositori();
                
                //_origRng.Select();
                //_origRng.Value = _origVal;

                // scrivere funzione che cancella tutte le modifiche ripristinate
                //Handler.StoreEdit(_origRng, tableName: MODIFICA);

                //ripristina commenti
                /*
                foreach (Excel.Range r in _origRng.Cells)
                {
                    if (!_comments[i])
                        r.ClearComments();

                    if (is_price)
                    {
                        _ws.Cells[r.Row - 2, r.Column].Value = _labelVenAcq[i];
                    }
                    else if (is_quantity)
                    {
                        _ws.Cells[r.Row - 1, r.Column].Value = _labelVenAcq[i];
                    }

                    i++;
                }
                is_Name_Match_Price(name_col_row_selected);
                */
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
                string suffData = Date.GetSuffissoData(new DateTime(int.Parse(data.Substring(0, 4)), int.Parse(data.Substring(4, 2)), int.Parse(data.Substring(6, 2)) ));
                string suffOra = Date.GetSuffissoOra(int.Parse(data.Substring(8, 2)) );
                int col = _definedNames.GetColFromDate(suffData, suffOra);
                int row = _definedNames.GetRowByName(dr["SiglaEntita"].ToString() + Simboli.UNION[0] + dr["SiglaInformazione"].ToString());

                _ws.Cells[row, col].Value2 = dr["Valore"].ToString() == "0" ? "" : dr["Valore"];

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
        }

        private void FormIncremento_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Workbook.Repository[MODIFICA].Rows.Count > 0)
            {
                //02/03/2017 FIX tolta richiesta di salvataggio.
                //DialogResult dr = MessageBox.Show("Salvare le modifiche apportate ai valori? Premere sì per inviare le modifiche al server, no per cancellarle definitivamente. Non sarà possibile recuperarle.", Simboli.NomeApplicazione + " - ATTENZIONE!!!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);

                //if (dr == DialogResult.Cancel)
                //{
                //    e.Cancel = true;
                //    return;
                //}

                //if (dr == DialogResult.Yes)
                    DataBase.SalvaModificheDB(MODIFICA);

                //if (dr == DialogResult.No)
                //    RipristinaValoriDaRepositori();
            }

            Workbook.Repository.Remove(MODIFICA);
            Workbook.Repository.Remove(ALL_OLD_VALUE);
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            Close();
        }


        // TO CHANGE - non è elegantissimo stabilire il questo modo il tipo di celle selezionate
        private bool is_Name_Match_Price(string input)
        {
            string pattern = @"^Offerta(.*)P$";
            string[] splitted = input.Split('.');
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
            return rgx.IsMatch(splitted[splitted.Length - 1]);
        }

        private bool is_Name_Match_Quantity(string input)
        {
            string pattern = @"^Offerta(.*)E$";
            string[] splitted = input.Split('.');
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
            return rgx.IsMatch(splitted[splitted.Length - 1]);
        }


        #endregion

        /*
        private void OnGoToSelectionChange(object sender, EventArgs e)
        {
            //abilito il panel con la textBox per l'inserimento del valore
            if (comboBoxVaiA.Text.Equals("PJOLLY"))
            {
                panelJolly.Visible = true;
            }
            else
            {
                panelJolly.Visible = false;
            }

            StateChanged_enableButton();
        }
        */

        // Gestisco l'abilitazione del bottone di modifica
        private void StateChanged_enableButton()
        {

            btnApplica.Enabled = false;

            if (!_selectionIsCorrect)
            {
                return;
            }

            // diversifico secondo il tipo di celle selezionate
            if (is_price)
            {
                bool textIsNull = ((this.txtPercentuale.Text == "" && this.rdbPercentuale.Checked)||(this.txtValore.Text == "" && this.rdbIncremento.Checked));
               if (!_valuesAreCorrect && !textIsNull )
                {
                    btnApplica.Enabled = false;
                    return;
                }

                if (comboBox_applicaA.SelectedIndex > -1)
                    btnApplica.Enabled = true;
            }
            else if (is_quantity)
            {
                if (comboBox_VaiA.SelectedIndex > -1)
                    btnApplica.Enabled = true;               
            }
        }

        private void StateChanged(object sender, EventArgs e)
        {
            StateChanged_enableButton();
        }
    }
}
