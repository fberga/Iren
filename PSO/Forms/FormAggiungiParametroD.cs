using Iren.ToolsExcel.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Iren.ToolsExcel.Forms
{
    public partial class FormAggiungiParametroD : Form
    {
        int _idTipologiaParametro;
        int _idApplicazione;
        int _idEntita;
        DataTable _valoriParametro;
        List<DateTime> _dateParNoNAttivi = new List<DateTime>();
        DateTime _dataFV = new DateTime(9999, 12, 31);
        bool _dataValidata = true;

        public FormAggiungiParametroD(string parametro, int idApplicazione, int idEntita, int idTipologiaParametro, DataTable valoriParametro)
        {
            InitializeComponent();
            Text = Simboli.nomeApplicazione + " - Aggiungi parametro giornaliero";

            groupDatiParD.Text = parametro;

            _idApplicazione = idApplicazione;
            _idTipologiaParametro = idTipologiaParametro;
            _idEntita = idEntita;
            _valoriParametro = valoriParametro;

            //in base ai valori già esistenti setto le date

            _dateParNoNAttivi =
                (from r in _valoriParametro.AsEnumerable()
                 where ((DateTime)r["Inizio Validità"]) > DateTime.Today
                 select (DateTime)r["Inizio Validità"]).ToList();
            _dateParNoNAttivi.Sort();
            
            dateTimeIV.MinDate = DateTime.Today.AddDays(1);
            dateTimeIV.Value = new DateTime(Math.Max(_dateParNoNAttivi.Last().Ticks, DateTime.Today.AddDays(1).Ticks));
        }

        private void btnAggiungi_Click(object sender, EventArgs e)
        {
            decimal value;

            if(!_dataValidata)
            {
                MessageBox.Show("Non può essere inserito un nuovo valore per la data selezionata. Controllare che non ci sia un altro valore con la stessa data di inizio validità o che non ci siano conflitti con altri valori inseriti.", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dateTimeIV.Focus();
                return;
            }

            if (!decimal.TryParse(txtValore.Text, NumberStyles.AllowDecimalPoint, CultureInfo.InstalledUICulture, out value))
            {
                MessageBox.Show("Il valore inserito non è valido.", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtValore.Focus();
                return;
            }

            //Utility.DataBase.Insert(Utility.DataBase.SP.INSERT_PARAMETRO, new Core.QryParams()
            //    {
            //        {"@IdEntita", _idEntita},
            //        {"@IdTipologiaParametro", _idTipologiaParametro},
            //        {"@DataIV", _oldDataIV.ToString("yyyyMMdd")},
            //        {"@DataFV", dateTimeIV.Value.AddDays(-1).ToString("yyyyMMdd")},
            //        {"@Valore", _oldValue},
            //        {"@Dettaglio", "D"}
            //    }
            //);

                //if (Utility.DataBase.Insert(Utility.DataBase.SP.INSERT_PARAMETRO, new Core.QryParams()
                //    {
                //        {"@IdEntita", _idEntita},
                //        {"@IdTipologiaParametro", _idTipologiaParametro},
                //        {"@DataIV", _oldDataIV.ToString("yyyyMMdd")},
                //        {"@DataFV", dateTimeIV.Value.AddDays(-1).ToString("yyyyMMdd")},
                //        {"@Valore", _oldValue},
                //        {"@Dettaglio", "D"}
                //    }))
                //{
                //    if (Utility.DataBase.Insert(Utility.DataBase.SP.UPDATE_PARAMETRO, new Core.QryParams()
                //        {
                //            {"@IdEntita", _idEntita},
                //            {"@IdTipologiaParametro", _idTipologiaParametro},
                //            {"@DataIV", dateTimeIV.Value.ToString("yyyyMMdd")},
                //            {"@DataFV", "99991231"},
                //            {"@Valore", value},
                //            {"@Dettaglio", "D"}
                //        }))
                //    {
                //        MessageBox.Show("Parametro aggiunto!", Simboli.nomeApplicazione + " - Aggiungi parametro", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                //    }
                //    else
                //    {
                //        //rollback della modifica fatta prima
                //        Utility.DataBase.Insert(Utility.DataBase.SP.UPDATE_PARAMETRO, new Core.QryParams()
                //        {
                //            {"@IdEntita", _idEntita},
                //            {"@IdTipologiaParametro", _idTipologiaParametro},
                //            {"@DataIV", _oldDataIV.ToString("yyyyMMdd")},
                //            {"@DataFV", "99991231"},
                //            {"@Valore", _oldValue},
                //            {"@Dettaglio", "D"}
                //        });
                //    }
                //}
                //else
                //{
                //    MessageBox.Show("Ci sono stati problemi nel salvataggio della modifica... Riprovare più tardi.", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //}        
        }

        private void txtValore_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
                e.Handled = false;
            }
        }

        private void txtValore_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
                e.SuppressKeyPress = true;
        }

        private void dateTimeIV_ValueChanged(object sender, EventArgs e)
        {
            _dataValidata = false;
        }

        private void dateTimeIV_Validating(object sender, CancelEventArgs e)
        {
            if (_dateParNoNAttivi.Contains(dateTimeIV.Value))
            {
                MessageBox.Show("Esiste già un valore con questa data di inizio validità: cancellarlo o cambiare la data per questo valore.", Simboli.nomeApplicazione + " - ATTENZIONE!!!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                _dataValidata = false;
                e.Cancel = true;
                
                return;
            }

            if(_dateParNoNAttivi.Max() > dateTimeIV.Value) 
            {
                DateTime tmpDataFV = _dateParNoNAttivi.First(dt => dt > dateTimeIV.Value).AddDays(-1);

                var result = MessageBox.Show("Per questo parametro esistono valori con data inizio validità superiore a quella scelta. La data fine validità sarà vincolata a '" + tmpDataFV.ToShortDateString() + "'. Continuare?", Simboli.nomeApplicazione + " - ATTENZIONE!!!", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);

                if (result == System.Windows.Forms.DialogResult.Cancel)
                {
                    _dataValidata = false;
                    e.Cancel = true;
                    return;
                }

                _dataFV = tmpDataFV;
            }

            _dataValidata = true;
        }
    }
}
