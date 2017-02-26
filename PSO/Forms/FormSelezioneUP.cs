using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Iren.PSO.Forms
{
    public partial class FormSelezioneUP : Form
    {
        #region Variabili
        
        private string _siglaInformazione = "";
        private Dictionary<string, string> _upList = new Dictionary<string, string>();
        
        #endregion

        #region Proprietà

        public List<string> ListaUP
        {
            get 
            {
                return _upList.Keys.ToList();
            }
        }
        
        #endregion

        #region Costruttori

        public FormSelezioneUP(string siglaInformazione)
        {
            InitializeComponent();
            _siglaInformazione = siglaInformazione;
            this.Text = Simboli.NomeApplicazione + " - Selezione UP";

            DataView entitaInformazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
            entitaInformazioni.RowFilter = "SiglaInformazione = '" + _siglaInformazione + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            string rowFilter = "SiglaEntita IN (";
            foreach (DataRowView entitaInfo in entitaInformazioni)
            {
                rowFilter += "'" + entitaInfo["SiglaEntita"] + "',";
            }
            rowFilter = rowFilter.Substring(0, rowFilter.Length - 1) + ")";

            DataView categorieEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categorieEntita.RowFilter = rowFilter + " AND IdApplicazione = " + Workbook.IdApplicazione;

            _upList =
                (from r in categorieEntita.ToTable(true, "SiglaEntita", "DesEntita").AsEnumerable()
                 select r).ToDictionary(r => r["SiglaEntita"].ToString(), r => r["DesEntita"].ToString());

            comboUP.DataSource = new BindingSource(_upList, null);
            comboUP.DisplayMember = "Value";
            comboUP.ValueMember = "Key";
            comboUP.SelectedIndex = 0;
        }

        #endregion

        #region Eventi

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            comboUP.SelectedIndex = -1;
            this.Close();
        }

        private void btnCarica_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Sposta la selezione sul titolo dell'UP scelta e ritorna la sua sigla.
        /// </summary>
        /// <returns>Restituisce la sigla dell'UP scelta.</returns>
        public new object ShowDialog()
        {
            base.ShowDialog();

            if (comboUP.SelectedIndex != -1) 
            {
                //non mi serve il nome del foglio perché lavoro direttamente con la siglaEntita
                DefinedNames n = new DefinedNames("", DefinedNames.InitType.GOTOs);
                string address = n.GetGotoFromSiglaEntita(comboUP.SelectedValue);
                Handler.Goto(address);
            }

            return comboUP.SelectedValue;
        } 

        #endregion
    }
}
