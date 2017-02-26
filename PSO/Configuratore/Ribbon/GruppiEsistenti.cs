using Iren.PSO.Base;
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class GruppiEsistenti : Form
    {
        DataTable _allGroups;
        DataTable _allFunctions;
        DataTable _utenti;
        Control _ribbon;

        public GruppiEsistenti(Control ribbon)
        {
            InitializeComponent();
            _ribbon = ribbon;

            _allGroups = DataBase.Select(DataBase.SP.RIBBON.GRUPPO_CONTROLLO, "@IdApplicazione=-1;@IdUtente=-1");
            _utenti = DataBase.Select(DataBase.SP.UTENTE_GRUPPO, "@IdUtenteGruppo=5");

            var unusedGroups = _allGroups.AsEnumerable()
                .Where(r => !ConfiguratoreRibbon.GruppoControlloUtilizzati.Contains((int)r["IdGruppoControllo"]))
                .Where(r => !ConfiguratoreRibbon.GruppiUtilizzati.Contains((int)r["IdGruppo"]))
                .Select(r => new { LabelGruppo = r["LabelGruppo"], IdGruppo = r["IdGruppo"] })
                .Distinct()                
                .ToList();

            listBoxGruppi.DisplayMember = "LabelGruppo";
            listBoxGruppi.ValueMember = "IdGruppo";

            _allFunctions = DataBase.Select(DataBase.SP.RIBBON.CONTROLLO_FUNZIONE);
            DataView funzioni = _allFunctions.DefaultView;

            funzioni.RowFilter = "IdFunzione=-1";
            listBoxFunzioni.DisplayMember = "NomeFunzione";
            listBoxFunzioni.ValueMember = "IdFunzione";

            listBoxFunzioni.DataSource = funzioni;
            listBoxGruppi.DataSource = unusedGroups;
        }

        private void CambioGruppo(object sender, EventArgs e)
        {
            if (listBoxGruppi.SelectedValue != null)
            {

                var users =
                    (from r in _allGroups.AsEnumerable()
                     join r1 in _utenti.AsEnumerable() on r["IdUtente"] equals r1["IdUtente"]
                     select new { IdUtente = r["IdUtente"], Nome = r1["Nome"] })
                .Distinct()
                .ToList();

                listBoxUtenti.ValueMember = "IdUtente";
                listBoxUtenti.DisplayMember = "Nome";
                listBoxUtenti.DataSource = users;
            }
        }

        private void CambioApplicazione(object sender, EventArgs e)
        {
            if (listBoxApplicazioni.SelectedValue != null)
            {
                //carico anteprima gruppo
                var controls = _allGroups.AsEnumerable()
                    .Where(r => r["IdGruppo"].Equals(listBoxGruppi.SelectedValue) && r["IdApplicazione"].Equals(listBoxApplicazioni.SelectedValue) && r["IdUtente"].Equals(listBoxUtenti.SelectedValue))
                    .ToList();

                RibbonGroup grp = new RibbonGroup(panelRibbonLayout, (int)listBoxGruppi.SelectedValue);
                grp.Text = listBoxGruppi.Text;
                panelRibbonLayout.Controls.Clear();
                panelRibbonLayout.Controls.Add(grp);

                foreach (DataRow r in controls)
                {
                    Control ctrl = Utility.AddControlToGroup(grp, r, _allFunctions);
                    ctrl.GotFocus += EvidenziaFunzioni;
                    ctrl.Tag = r["IdGruppoControllo"];
                }
            }
        }

        private void CambioUtente(object sender, EventArgs e)
        {
            if (listBoxUtenti.SelectedValue != null)
            {
                var applications = _allGroups.AsEnumerable()
                    .Where(r => r["IdGruppo"].Equals(listBoxGruppi.SelectedValue) && r["IdUtente"].Equals(listBoxUtenti.SelectedValue))
                    .Select(r => new { IdApplicazione = r["IdApplicazione"], DesApplicazione = r["DesApplicazione"] })
                    .OrderBy(r => r.IdApplicazione)
                    .Distinct()
                    .ToList();

                listBoxApplicazioni.DataSource = applications;
                listBoxApplicazioni.ValueMember = "IdApplicazione";
                listBoxApplicazioni.DisplayMember = "DesApplicazione";
            }
        }

        private void EvidenziaFunzioni(object sender, EventArgs e)
        {
            IRibbonControl ctrl = sender as IRibbonControl;
            if (ctrl.Functions.Count > 0)
                _allFunctions.DefaultView.RowFilter = "IdGruppoControllo = " + ((Control)ctrl).Tag + " AND IdFunzione IN (" + string.Join(",", ctrl.Functions) + ")";
            else
                _allFunctions.DefaultView.RowFilter = "IdFunzione = -1";
        }

        private void AggiungiGruppo_Click(object sender, EventArgs e)
        {
            var ribbonGroup = panelRibbonLayout.Controls.OfType<RibbonGroup>().First();
            var ctrls = Utility.GetAll(ribbonGroup);

            foreach (Control ctrl in ctrls)
                ctrl.GotFocus -= EvidenziaFunzioni;

            Utility.AddGroupToRibbon(_ribbon, ribbonGroup);
            
            //listBoxGruppi.Items.RemoveAt(listBoxGruppi.SelectedIndex);
        }
    }
}
