using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class CopiaConfigurazione : Form
    {
        DataTable _gruppoControllo;
        DataTable _applicazioni;
        DataTable _utenti;


        
        public CopiaConfigurazione()
        {
            InitializeComponent();

            _gruppoControllo = DataBase.Select(DataBase.SP.RIBBON.GRUPPO_CONTROLLO, "@IdApplicazione=-1;@IdUtente=-1");
            _applicazioni = DataBase.Select(DataBase.SP.APPLICAZIONE, "@IdApplicazione=0");
            _utenti = DataBase.Select(DataBase.SP.UTENTE_GRUPPO, "@IdUtenteGruppo=0");

            _utenti.DefaultView.RowFilter = "IdUtenteGruppo = 1 OR IdUtenteGruppo = 5";

            var utentiFrom = 
                (from r in _gruppoControllo.AsEnumerable()
                 join r1 in _utenti.AsEnumerable() on r["IdUtente"] equals r1["IdUtente"]
                 where r1["IdUtenteGruppo"].Equals(1) || r1["IdUtenteGruppo"].Equals(5)
                 orderby r["IdUtente"]
                 select new KeyValuePair<int, string>((int)r["IdUtente"], r1["Nome"].ToString()))
                 .Distinct()
                 .ToList();

            listBoxUtentiTo.ValueMember = "IdUtente";
            listBoxUtentiTo.DisplayMember = "Nome";
            listBoxUtentiTo.DataSource = _utenti.DefaultView;

            listBoxUtentiFrom.ValueMember = "Key";
            listBoxUtentiFrom.DisplayMember = "Value";
            listBoxUtentiFrom.DataSource = utentiFrom;
        }

        private void listBoxUtentiFrom_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxUtentiFrom.SelectedIndex != -1)
            {
                var applicazioni =
                    (from r in _gruppoControllo.AsEnumerable()
                     join r1 in _applicazioni.AsEnumerable() on r["IdApplicazione"] equals r1["IdApplicazione"]
                     where r["IdUtente"].Equals(listBoxUtentiFrom.SelectedValue)
                     orderby r["DesApplicazione"]
                     select new KeyValuePair<int, string>((int)r["IdApplicazione"], r["DesApplicazione"].ToString()))
                     .Distinct()
                     .ToList();
                    

                listBoxApplicazioni.DataSource = applicazioni;
                listBoxApplicazioni.ValueMember = "Key";
                listBoxApplicazioni.DisplayMember = "Value";

                ((DataView)listBoxUtentiTo.DataSource).RowFilter = "IdUtenteGruppo = 1 OR IdUtenteGruppo = 5 AND IdUtente <> " + listBoxUtentiFrom.SelectedValue;

            }
            else
            {
                listBoxApplicazioni.DataSource = null;
                ((DataView)listBoxUtentiTo.DataSource).RowFilter = "";
            }
        }

        private void btnEmpty_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listBoxApplicazioni.Items.Count; i++)
                listBoxApplicazioni.SetSelected(i, false);
        }

        private void btnCopia_Click(object sender, EventArgs e)
        {
            if (listBoxUtentiTo.SelectedItems.Count > 0)
            {
                if (listBoxApplicazioni.SelectedItems.Count > 0)
                {
                    foreach (KeyValuePair<int,string> applicazione in listBoxApplicazioni.SelectedItems)
                        foreach (DataRowView utenteTo in listBoxUtentiTo.SelectedItems)
                            DataBase.Insert(DataBase.SP.RIBBON.COPIA_CONFIGURAZIONE, new Iren.PSO.Core.QryParams() 
                            {
                                {"@IdApplicazione", applicazione.Key }, 
                                {"@IdUtenteFrom", listBoxUtentiFrom.SelectedValue}, 
                                {"@IdUtenteTo", utenteTo["IdUtente"]}
                            });
                }
                else
                {
                    foreach (KeyValuePair<int,string> applicazione in listBoxApplicazioni.Items)
                        foreach (DataRowView utenteTo in listBoxUtentiTo.SelectedItems)
                            DataBase.Insert(DataBase.SP.RIBBON.COPIA_CONFIGURAZIONE, new Iren.PSO.Core.QryParams() 
                            {
                                {"@IdApplicazione", applicazione.Key }, 
                                {"@IdUtenteFrom", listBoxUtentiFrom.SelectedValue}, 
                                {"@IdUtenteTo", utenteTo["IdUtente"]}
                            });
                }
            }
        }
    }
}
