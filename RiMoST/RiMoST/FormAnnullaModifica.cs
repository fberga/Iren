using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.PSO.Core;

namespace Iren.RiMoST
{
    public partial class FormAnnullaModifica : Form
    {
        #region Variabili
        string _anno;
        #endregion

        #region Costruttori
        public FormAnnullaModifica(string anno)
        {
            _anno = anno;
            InitializeComponent();
        }
        #endregion

        #region Callbacks
        private void FormAnnullaModifica_Load(object sender, EventArgs e)
        {
            if(ThisDocument.DB.OpenConnection())
            {
                DataView dv = (ThisDocument.DB.Select("spGetRichiesta", "@IdStruttura=" + ThisDocument._idStruttura) ?? new DataTable()).DefaultView;
                dv.RowFilter = "IdTipologiaStato NOT IN (4, 7) AND IdRichiesta LIKE '%" + _anno + "'";
                cmbRichiesta.DataSource = dv;
                cmbRichiesta.DisplayMember = "IdRichiesta";

                ThisDocument.DB.CloseConnection();
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Sei sicuro di voler ANNULLARE la richiesta selezionata?", "Attenzione!", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.OK)
            {
                DataRowView row = (DataRowView)cmbRichiesta.SelectedItem;
                try
                {
                    if (ThisDocument.DB.OpenConnection())
                    {
                        ThisDocument.DB.Insert("spAnnullaRichiesta", new QryParams() 
                        {
                            {"@IdRichiesta", row["IdRichiesta"]},
                            {"@IdStruttura", ThisDocument._idStruttura},
                        });
                        ThisDocument.DB.CloseConnection();
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Errore nell'annullamento della richiesta. Riporvare più tardi.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            this.Close();
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmbRichiesta_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataRowView row = (DataRowView)cmbRichiesta.SelectedItem;
            string path = @"file:///" + row["NomeFile"];
            DocPreview.Navigate(path);
        }
        #endregion
    }
}
