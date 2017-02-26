using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.PSO.Core;
using System.Globalization;
using System.Configuration;

namespace Iren.RiMoST
{
    public partial class SelezionaModifica : Form
    {
        #region Variabili
        string _anno;
        public bool _chkIsDraft;
        public bool _btnRefreshEnabled;
        #endregion

        #region Costruttori
        public SelezionaModifica(string anno, bool chkIsDraft, bool btnRefreshEnabled)
        {
            _anno = anno;
            _chkIsDraft = chkIsDraft;
            _btnRefreshEnabled = btnRefreshEnabled;
            InitializeComponent();
        }
        #endregion

        #region Callbacks
        private void SelezionaModifica_Load(object sender, EventArgs e)
        {
            if(ThisDocument.DB.OpenConnection())
            {
                DataView dv = (ThisDocument.DB.Select("spGetRichiesta", "@IdTipologiaStato:7; @IdStruttura:" + ThisDocument._idStruttura) ?? new DataTable()).DefaultView;
                dv.RowFilter = "IdRichiesta LIKE '%" + _anno + "'";
                cmbRichiesta.DataSource = dv;
                cmbRichiesta.DisplayMember = "IdRichiesta";

                ThisDocument.DB.CloseConnection();
            }
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (cmbRichiesta.Text == "")
            {
                MessageBox.Show("Non ci sono bozze al momento.", "Attenzione!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                DataRowView row = (DataRowView)cmbRichiesta.SelectedItem;
                Globals.ThisDocument.lbIdRichiesta.LockContents = false;
                Globals.ThisDocument.lbIdRichiesta.Text = row["IdRichiesta"].ToString();
                Globals.ThisDocument.lbIdRichiesta.LockContents = true;
                //Globals.ThisDocument.dropDownStrumenti.DropDownListEntries.Add(row["DesApplicazione"].ToString(), row["IdApplicazione"].ToString());
                DataView applicazioni = new DataView(Globals.ThisDocument.Applicazioni);
                applicazioni.RowFilter = "IdApplicazione = " + row["IdApplicazione"];
                //Globals.ThisDocument.dropDownStrumenti.DropDownListEntries.Add(applicazioni[0]["DesApplicazione"].ToString(), row["IdApplicazione"].ToString());


                int index = Globals.ThisDocument.dropDownStrumenti.DropDownListEntries.OfType<Microsoft.Office.Interop.Word.ContentControlListEntry>().First(c => c.Text == applicazioni[0]["DesApplicazione"].ToString()).Index;
                Globals.ThisDocument.dropDownStrumenti.DropDownListEntries[index].Select();

                Globals.ThisDocument.dropDownStrumenti.LockContents = true;

                
                //Globals.ThisDocument.dropDownStrumenti.DropDownListEntries[1].Select();                

                Globals.ThisDocument.txtOggetto.Text = "" + row["Oggetto"];
                Globals.ThisDocument.txtDescrizione.Text = "" + row["Descr"];
                Globals.ThisDocument.txtNote.Text = "" + row["Note"];
                
                _chkIsDraft = true;
                _btnRefreshEnabled = false;
                
                this.Hide();
            }
        }
        #endregion
    }
}
