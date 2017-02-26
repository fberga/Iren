using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Forms
{
    public partial class FormIncremento : Form
    {
        public const string MODIFICA = "FormIncrementoModifica";

        #region Variabili
        private object[,] _origVal;

        private Excel.Range _origRng;

        private DefinedNames _definedNames;

        private Excel.Worksheet _ws;

        private double? _percentage;
        private double? _increment;

        private bool _selectionIsCorrect = false;
        private bool _valuesAreCorrect = false;

        #endregion

        #region Costruttore

        public FormIncremento(Excel.Worksheet ws, Excel.Range rng)
        {
            InitializeComponent();
            this.Text = Simboli.NomeApplicazione + " - Incremento";

            _ws = ws;

            _ws.SelectionChange += ChangeSelectionToIncrement;
            _definedNames = new DefinedNames(_ws.Name, DefinedNames.InitType.All);

            Workbook.Repository.Add(Workbook.Repository.CreaTabellaModifica(MODIFICA));

            btnRipristina.Enabled = false;
            btnApplica.Enabled = false;
            ChangeSelectionToIncrement(Workbook.Application.Selection);
        }

        #endregion

        #region Eventi

        private void ChangeSelectionToIncrement(Excel.Range Target)
        {
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

            int marketOffset = Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1") ? Simboli.GetMarketOffset(DateTime.Now.Hour) : 0;
            int firstCol = _definedNames.GetColFromDate(Date.SuffissoDATA1);

            if (Target.Column < firstCol + marketOffset)
            {
                lbErrore.Text = "ERRORE: Il range selezionato contiene celle appartenenti a mercati chiusi.";
                return;
            }

            _selectionIsCorrect = true;

            if (_valuesAreCorrect)
                btnApplica.Enabled = true;

            btnRipristina.Enabled = false;

            if (Target.Cells.Count == 1)
            {
                _origVal = new object[1, 1];
                _origVal[0, 0] = Target.Value;
            }
            else
                _origVal = Target.Value;
            
            _origRng = Target;

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
                _valuesAreCorrect = false;
                btnApplica.Enabled = false;
                return;
            }

            double val;
            string text = txt.Text.Replace(".", ",");

            if (Double.TryParse(text, out val))
            {
                btnApplica.Enabled = true;
            }
            else
            {
                _valuesAreCorrect = false;
                btnApplica.Enabled = false;
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

            if (!_selectionIsCorrect)
                btnApplica.Enabled = false;
        }

        private void TextElements_EnabledChanged(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;
            if (txt.Enabled == true)
                TextElements_TextChanged(sender, e);
        }

        private void btnApplica_Click(object sender, EventArgs e)
        {
            Sheet.Protected = false;

            foreach (Excel.Range rng in _origRng.Cells)
            {
                if (rng.Value != null)
                {
                    double val = (double)rng.Value;
                    if (_percentage != null)
                    {
                        rng.Value = val + val * (_percentage.Value/100);
                    }
                    else if (_increment != null)
                    {
                        rng.Value += _increment.Value;
                    }
                }
            }

            Handler.StoreEdit(_origRng, tableName: MODIFICA);

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
            if (MessageBox.Show("Ripristinare i valori originali? Premere sì per continuare, no per lasciare i valori attuali.", Simboli.NomeApplicazione + " - ATTENZIONE!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                Sheet.Protected = false;
                _origRng.Select();
                _origRng.Value = _origVal;

                Handler.StoreEdit(_origRng, tableName: MODIFICA);

                Sheet.Protected = true;
            }
        }

        private void FormIncremento_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Workbook.Repository[MODIFICA].Rows.Count > 0)
            {
                DialogResult dr = MessageBox.Show("Salvare le modifiche apportate ai valori? Premere sì per inviare le modifiche al server, no per cancellarle definitivamente. Non sarà possibile recuperarle.", Simboli.NomeApplicazione + " - ATTENZIONE!!!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);

                if (dr == DialogResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }

                if (dr == DialogResult.Yes)
                    DataBase.SalvaModificheDB(MODIFICA);
            }

            Workbook.Repository.Remove(MODIFICA);
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            Close();
        }

        #endregion
    }
}
