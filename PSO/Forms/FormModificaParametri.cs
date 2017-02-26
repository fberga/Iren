using Iren.PSO.Base;
using System;
using System.Data;
using System.Globalization;
using System.Windows.Forms;


namespace Iren.PSO.Forms
{
    public partial class FormModificaParametri : Form
    {
        #region Variabili
        
        DataView _dvParametri = new DataView();
        DataTable _parametri;
        DataView _entita;
        bool _onEdit = false;
        bool _isUpdate = false;
        DateTime _currDataIV;
        
        #endregion

        #region Costanti
        
        const string 
            INIZIO_VALIDITA = "Inizio validità",
            FINE_VALIDITA = "Fine validità",
            VALORE = "Valore";

        #endregion

        #region Costruttore

        public FormModificaParametri() 
        {
            InitializeComponent();

            this.Text = Simboli.NomeApplicazione + " - Modifica Parametri";

            _entita = new DataView(Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA]);
            if(DataBase.OpenConnection())
            {
                _parametri = DataBase.Select(DataBase.SP.PAR.ELENCO_PARAMETRI, "@IdApplicazione=" + Workbook.IdApplicazione) ?? new DataTable();

                _dvParametri = new DataView(_parametri);

                cmbEntita.ValueMember = "SiglaEntita";
                cmbEntita.DisplayMember = "DesEntita";

                cmbParametri.DisplayMember = "Descrizione";

                cmbEntita.DataSource = _entita;
                cmbParametri.DataSource = _dvParametri;

                dataGridParametri.CellBeginEdit += ParameterCellBeginEdit;
                
                dataGridParametri.CurrentCellChanged += CurrentCellChanged;
                dataGridParametri.RowDirtyStateNeeded += IsRowDirty;
            }
            else
            {
                MessageBox.Show("Non è possibile modificare i valori dei parametri in assenza di connessione...", Simboli.NomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        #endregion

        #region Metodi

        private void IsRowDirty(object sender, QuestionEventArgs e) 
        {
            e.Response = _onEdit;
        }
        private bool IsModEnabled(DataGridViewRow parameter) 
        {
            if(parameter.Cells[INIZIO_VALIDITA].Value != null && (DateTime)parameter.Cells[INIZIO_VALIDITA].Value >= DateTime.Today)
                return true;

            return false;
        }
        private bool IsInsertBeforeEnabled(DataTable parameters, int index) 
        {
            DateTime currentIV = (DateTime)parameters.Rows[index][INIZIO_VALIDITA];
            DateTime precedingIV = index == 0 ? DateTime.MinValue : (DateTime)parameters.Rows[index - 1][INIZIO_VALIDITA];
            DateTime precedingFV = index == 0 ? DateTime.MinValue : (DateTime)parameters.Rows[index - 1][FINE_VALIDITA];

            return
                currentIV > DateTime.Today                              //posso inserire un parametro con IV >= oggi
                && precedingFV > DateTime.Today.AddDays(-1)             //posso arretrare di 1 giorno la fine validità della riga sopra
                && precedingIV != precedingFV;                          //ho spazio per ridimensionare la fine validità della riga sopra
        }
        private bool IsInsertAfterEnabled(DataTable parameters, int index) 
        {
            if (index == parameters.Rows.Count - 1)
                return true;

            DateTime currentIV = (DateTime)parameters.Rows[index][INIZIO_VALIDITA];
            DateTime currentFV = (DateTime)parameters.Rows[index][FINE_VALIDITA];
            DateTime subsequentIV = (DateTime)parameters.Rows[index + 1][INIZIO_VALIDITA];

            return
                subsequentIV > DateTime.Today                           //posso inserire un parametro con IV >= oggi
                && currentFV > DateTime.Today.AddDays(-1)               //posso arretrare di 1 giorno la fine validità della riga corrente
                && currentIV < currentFV;                               //ho spazio per ridimensionare la fine validità della riga corrente
        }

        private void RefreshMenuItems() 
        {
            if (!_onEdit)
            {
                if (dataGridParametri.CurrentRow != null || dataGridParametri.IsCurrentRowDirty)
                {
                    if (IsModEnabled(dataGridParametri.CurrentRow))
                    {
                        modificareValoreContextMenu.Enabled = true;
                        cancellaParametroContextMenu.Enabled = true;
                        modificaTopMenu.Enabled = true;
                        elimiaTopMenu.Enabled = true;
                    }
                    else
                    {
                        modificareValoreContextMenu.Enabled = false;
                        cancellaParametroContextMenu.Enabled = false;
                        modificaTopMenu.Enabled = false;
                        elimiaTopMenu.Enabled = false;
                    }

                    if (IsInsertAfterEnabled((DataTable)dataGridParametri.DataSource, dataGridParametri.CurrentRow.Index))
                    {
                        inserisciSottoContextMenu.Enabled = true;
                        inserisciSottoTopMenu.Enabled = true;
                    }
                    else
                    {
                        inserisciSottoContextMenu.Enabled = false;
                        inserisciSottoTopMenu.Enabled = false;
                    }

                    if (IsInsertBeforeEnabled((DataTable)dataGridParametri.DataSource, dataGridParametri.CurrentRow.Index))
                    {
                        inserisciSopraContextMenu.Enabled = true;
                        inserisciSopraTopMenu.Enabled = true;
                    }
                    else
                    {
                        inserisciSopraContextMenu.Enabled = false;
                        inserisciSopraTopMenu.Enabled = false;
                    }
                }
                else
                {
                    inserisciSopraContextMenu.Enabled = false;
                    inserisciSopraTopMenu.Enabled = false;
                    inserisciSottoContextMenu.Enabled = false;
                    inserisciSottoTopMenu.Enabled = false;
                    modificareValoreContextMenu.Enabled = false;
                    cancellaParametroContextMenu.Enabled = false;
                    modificaTopMenu.Enabled = false;
                    elimiaTopMenu.Enabled = false;
                }
            }
        }

        #endregion

        #region Eventi

        private void CurrentCellChanged(object sender, EventArgs e) 
        {
            if (_onEdit)
            {
                dataGridParametri.BeginEdit(false);
            }
            else
            {
                RefreshMenuItems();
            }
        }
        private void ParameterCellBeginEdit(object sender, DataGridViewCellCancelEventArgs e) 
        {
            if (e.ColumnIndex != 0 && e.ColumnIndex != 2)
            {
                e.Cancel = true;
                return;
            }

            //evito di assegnare più eventi se cambio la cella nella stessa riga
            if (!_onEdit)
            {
                _onEdit = true;
                modificareValoreContextMenu.Enabled = false;
                cancellaParametroContextMenu.Enabled = false;
                modificaTopMenu.Enabled = false;
                elimiaTopMenu.Enabled = false;
                inserisciSopraContextMenu.Enabled = false;
                inserisciSopraTopMenu.Enabled = false;
                inserisciSottoTopMenu.Enabled = false;
                inserisciSottoContextMenu.Enabled = false;

                dataGridParametri.CellValidating += ParameterValidating;
                dataGridParametri.RowValidating += ParameterRowValidating;
                dataGridParametri.RowValidated += ParameterRowValidated;
            }
        }
        private void ParameterValidating(object sender, DataGridViewCellValidatingEventArgs e) 
        {
            if (dataGridParametri.IsCurrentCellDirty && dataGridParametri.EditingControl != null)
            {
                string value = e.FormattedValue.ToString();

                //salvo la vecchia dataIV per fare l'update
                _currDataIV = (DateTime)dataGridParametri[INIZIO_VALIDITA, e.RowIndex].Value;

                switch (dataGridParametri.Columns[e.ColumnIndex].Name)
                {
                    case INIZIO_VALIDITA:
                        DateTime date = new DateTime();

                        if (DateTime.TryParseExact(value, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out date) ||
                            DateTime.TryParseExact(value, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out date) ||
                            DateTime.TryParseExact(value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out date) ||
                            DateTime.TryParseExact(value, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out date) ||
                            value == "-")
                        {
                            dataGridParametri.EditingControl.Text = date.ToString("dd/MM/yyyy");

                            if (date < DateTime.Today)
                            {
                                MessageBox.Show("La data di inizio vaidità non può essere antecedente a oggi!", Simboli.NomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                e.Cancel = true;
                            }
                        }
                        else
                        {
                            e.Cancel = true;
                        }
                        break;
                    case VALORE:
                        double number;
                        if (value == "")
                        {
                            MessageBox.Show("Non è possibile lasciare il campo Valore vuoto!", Simboli.NomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            e.Cancel = true;
                        }
                        else if (double.TryParse(value, out number) ||
                            double.TryParse(value, NumberStyles.AllowDecimalPoint, CultureInfo.InstalledUICulture, out number))
                        {
                            dataGridParametri.EditingControl.Text = number.ToString(CultureInfo.InstalledUICulture);
                        }
                        else
                        {
                            MessageBox.Show("Il valore inserito non è un numero valido!", Simboli.NomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            e.Cancel = true;
                        }
                        break;
                }
            }
        }
        private void ParameterRowValidating(object sender, DataGridViewCellCancelEventArgs e) 
        {
            DateTime currentIV = (DateTime)dataGridParametri[INIZIO_VALIDITA, e.RowIndex].Value;

            //controllo, se esiste, la dataIV della riga successiva ed eventualmente aggiusto la fine validità
            if (e.RowIndex < dataGridParametri.Rows.Count - 1)
            {
                DateTime subsequentIV = (DateTime)dataGridParametri[INIZIO_VALIDITA, e.RowIndex + 1].Value;
                DateTime currentFV = dataGridParametri[FINE_VALIDITA, e.RowIndex].Value is DBNull ? subsequentIV : (DateTime)dataGridParametri[FINE_VALIDITA, e.RowIndex].Value;

                if (currentIV >= subsequentIV)
                {
                    MessageBox.Show("La data di inizio validità della riga corrente va in conflitto con quella della successiva.", Simboli.NomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                    return;
                }
                if (subsequentIV - currentFV != new TimeSpan(1, 0, 0, 0))
                    dataGridParametri[FINE_VALIDITA, e.RowIndex].Value = subsequentIV.AddDays(-1);
            }
            //controllo, se esiste, la dataFV della riga precedente ed eventualmente la aggiorno
            if (e.RowIndex > 0)
            {
                DateTime precedingIV = (DateTime)dataGridParametri[INIZIO_VALIDITA, e.RowIndex - 1].Value;
                if (dataGridParametri[FINE_VALIDITA, e.RowIndex - 1].Value is DBNull)
                {
                    dataGridParametri[FINE_VALIDITA, e.RowIndex - 1].Value = currentIV.AddDays(-1);
                }
                else
                {
                    DateTime precedingFV = (DateTime)dataGridParametri[FINE_VALIDITA, e.RowIndex - 1].Value;
                    if (currentIV - precedingIV < new TimeSpan(1, 0, 0, 0))
                    {
                        MessageBox.Show("La data di inizio validità della riga corrente va in conflitto con quella della precedente.", Simboli.NomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        e.Cancel = true;

                        return;
                    }

                    if (currentIV - precedingFV != new TimeSpan(1, 0, 0, 0))
                    {
                        dataGridParametri[FINE_VALIDITA, e.RowIndex - 1].Value = currentIV.AddDays(-1);
                    }
                }
            }

            //vedo se devo eliminare la data di fine validità
            if (e.RowIndex == dataGridParametri.Rows.Count - 1 && dataGridParametri[FINE_VALIDITA, e.RowIndex].Value != DBNull.Value)
            {
                dataGridParametri[FINE_VALIDITA, e.RowIndex].Value = DBNull.Value;
            }
        }
        private void ParameterRowValidated(object sender, DataGridViewCellEventArgs e) 
        {
            DataRowView parRow = (DataRowView)cmbParametri.SelectedValue;

            if (dataGridParametri.IsCurrentRowDirty)
            {
                DateTime precIV = (DateTime)dataGridParametri[INIZIO_VALIDITA, e.RowIndex - 1].Value;
                DateTime newIV = (DateTime)dataGridParametri[INIZIO_VALIDITA, e.RowIndex].Value;
                DateTime newFV = dataGridParametri[FINE_VALIDITA, e.RowIndex].Value is DBNull ? DateTime.MaxValue : (DateTime)dataGridParametri[FINE_VALIDITA, e.RowIndex].Value;
                double value = (double)dataGridParametri[VALORE, e.RowIndex].Value;
                
                //insert or update
                if (_isUpdate)
                {
                    DataBase.Insert(DataBase.SP.PAR.UPDATE_PARAMETRO, new Core.QryParams()
                    {
                        {"@IdEntita", parRow["IdEntita"]},
                        {"@IdTipologiaParametro", parRow["IdParametro"]},
                        {"@CurrDataIV", _currDataIV.ToString("yyyyMMdd")},
                        {"@NewDataIV", newIV.ToString("yyyyMMdd")},
                        {"@NewDataFV", newFV.ToString("yyyyMMdd")},
                        {"@Valore", value}
                    });


                }
                else
                {
                    DataBase.Insert(DataBase.SP.PAR.INSERT_PARAMETRO, new Core.QryParams()
                    {
                        {"@IdEntita", parRow["IdEntita"]},
                        {"@IdTipologiaParametro", parRow["IdParametro"]},
                        {"@PrecDataIV", precIV.ToString("yyyyMMdd")},
                        {"@NewDataIV", newIV.ToString("yyyyMMdd")},
                        {"@NewDataFV", newFV.ToString("yyyyMMdd")},
                        {"@Valore", value}
                    });
                }
            }

            _onEdit = false;

            dataGridParametri.RowValidated -= ParameterRowValidated;
            dataGridParametri.RowValidating -= ParameterRowValidating;
        }

        private void CambiaEntita(object sender, EventArgs e) 
        {
            if (_parametri.Columns.Contains("SiglaEntita"))
            {
                _dvParametri.RowFilter = "SiglaEntita = '" + cmbEntita.SelectedValue + "'";
                CambiaParametro(null, null);
            }
        }
        private void CambiaParametro(object sender, EventArgs e) 
        {
            if (cmbParametri.SelectedValue != null)
            {
                DataRowView r = cmbParametri.SelectedValue as DataRowView;

                DataTable valori = DataBase.Select(DataBase.SP.PAR.VALORI_PARAMETRI, new Core.QryParams() 
                    {
                        {"@IdApplicazione", Workbook.IdApplicazione},
                        {"@IdEntita", r["IdEntita"]},
                        {"@IdTipologiaParametro", r["IdParametro"]},
                    }) ?? new DataTable();

                DataTable valCorretti = new DataTable()
                {
                    Columns = 
                        {
                            {INIZIO_VALIDITA, typeof(DateTime)},
                            {FINE_VALIDITA, typeof(DateTime)},
                            {VALORE, typeof(double)}
                        }
                };


                foreach (DataRow val in valori.Rows)
                {
                    DataRow newRow = valCorretti.NewRow();

                    DateTime fineValidita = DateTime.ParseExact(val["DataFV"].ToString().Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);

                    newRow[INIZIO_VALIDITA] = DateTime.ParseExact(val["DataIV"].ToString().Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);
                    if (fineValidita.Year != 9999)
                        newRow[FINE_VALIDITA] = fineValidita;
                    newRow[VALORE] = decimal.Parse(val[VALORE].ToString(), CultureInfo.CurrentUICulture);

                    valCorretti.Rows.Add(newRow);
                }

                dataGridParametri.DataSource = valCorretti;
                dataGridParametri.Columns[INIZIO_VALIDITA].DefaultCellStyle.FormatProvider = CultureInfo.InstalledUICulture;
                dataGridParametri.Columns[FINE_VALIDITA].DefaultCellStyle.FormatProvider = CultureInfo.InstalledUICulture;
                dataGridParametri.Columns[FINE_VALIDITA].DefaultCellStyle.NullValue = "-";
                dataGridParametri.Columns[VALORE].DefaultCellStyle.Format = "0.#########";

                foreach (DataGridViewColumn c in dataGridParametri.Columns)
                    c.SortMode = DataGridViewColumnSortMode.NotSortable;

                if (dataGridParametri.Rows.Count > 0)
                {
                    dataGridParametri.CurrentCell = dataGridParametri[INIZIO_VALIDITA, dataGridParametri.Rows.Count - 1];
                }
            }
            else
            {
                dataGridParametri.DataSource = null;
            }
        }

        private void FormParametriClose(object sender, EventArgs e) 
        {
            this.Close();
        }

        private void ModificaParametro(object sender, EventArgs e) 
        {
            if (dataGridParametri.CurrentCell.OwningColumn.Name != INIZIO_VALIDITA ||
                dataGridParametri.CurrentCell.OwningColumn.Name != VALORE)
                dataGridParametri.CurrentCell = dataGridParametri[INIZIO_VALIDITA, dataGridParametri.CurrentCell.RowIndex];
            _isUpdate = true;
            dataGridParametri.BeginEdit(false);
        }
        private void InserisciSopra(object sender, EventArgs e) 
        {
            int index = dataGridParametri.CurrentRow.Index;

            DataTable dt = (DataTable)dataGridParametri.DataSource;

            DateTime precedingFV = (DateTime)dt.Rows[index - 1][FINE_VALIDITA];

            //inserisco la nuova riga
            DataRow r = dt.NewRow();
            r[INIZIO_VALIDITA] = precedingFV;
            r[FINE_VALIDITA] = precedingFV;
            dt.Rows.InsertAt(r, index);

            //metto la datagrid in modifica
            dataGridParametri.CurrentCellChanged -= CurrentCellChanged;
            dataGridParametri.CurrentCell = dataGridParametri[VALORE, index];
            dataGridParametri.CurrentCellChanged += CurrentCellChanged;
            _isUpdate = false;
            dataGridParametri.BeginEdit(false);
        }
        private void InserisciSotto(object sender, EventArgs e) 
        {
            int index = dataGridParametri.CurrentRow.Index;

            DataTable dt = (DataTable)dataGridParametri.DataSource;

            DateTime precedingFV = dt.Rows[index][FINE_VALIDITA] is DBNull ? DateTime.Today.AddDays(-1) : (DateTime)dt.Rows[index][FINE_VALIDITA];
            DateTime iv = precedingFV.AddDays(1);
            DateTime fv = dt.Rows[index][FINE_VALIDITA] is DBNull ? DateTime.MaxValue : ((DateTime)dt.Rows[index + 1][INIZIO_VALIDITA]).AddDays(-1);

            if (iv > fv)
                iv = fv;

            //inserisco la nuova riga
            DataRow r = dt.NewRow();
            r[INIZIO_VALIDITA] = iv;
            if (fv != DateTime.MaxValue)
                r[FINE_VALIDITA] = fv;
            dt.Rows.InsertAt(r, index + 1);

            //metto la datagrid in modifica
            dataGridParametri.CurrentCellChanged -= CurrentCellChanged;
            dataGridParametri.CurrentCell = dataGridParametri[VALORE, index + 1];
            dataGridParametri.CurrentCellChanged += CurrentCellChanged;
            _isUpdate = false;
            dataGridParametri.BeginEdit(false);
        }

        private void CancellaParametro(object sender, EventArgs e) 
        {
            if (MessageBox.Show("Eliminare la riga?", Simboli.NomeApplicazione, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == System.Windows.Forms.DialogResult.Yes)
            {
                DataRowView parRow = (DataRowView)cmbParametri.SelectedValue;

                DateTime dataIV = (DateTime)dataGridParametri[INIZIO_VALIDITA, dataGridParametri.CurrentRow.Index].Value;
                DateTime dataFV = dataGridParametri[FINE_VALIDITA, dataGridParametri.CurrentRow.Index].Value is DBNull ? DateTime.MaxValue : (DateTime)dataGridParametri[FINE_VALIDITA, dataGridParametri.CurrentRow.Index].Value;


                DataBase.Insert(DataBase.SP.PAR.DELETE_PARAMETRO, new Core.QryParams()
                        {
                            {"@IdEntita", parRow["IdEntita"]},
                            {"@IdTipologiaParametro", parRow["IdParametro"]},
                            {"@DataIV", dataIV.ToString("yyyyMMdd")},
                            {"@DataFV", dataFV.ToString("yyyyMMdd")},
                        });

                int index = cmbParametri.SelectedIndex;
                cmbParametri.SelectedIndex = -1;
                cmbParametri.SelectedIndex = index;
            }
        }

        #endregion
    }
}
