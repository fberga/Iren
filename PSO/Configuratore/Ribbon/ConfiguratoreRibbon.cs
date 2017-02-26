using Iren.PSO.Base;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class ConfiguratoreRibbon : Form
    {
        //private string[] _ambienti;

        //public static int IdApplicazione { get; private set; }
        public static List<int> IdUtenti { get { return new List<int>() { 62 }; } }// private set; }

        public static List<int> ControlliUtilizzati { get; private set; }
        public static List<int> GruppiUtilizzati { get; private set; }
        public static List<int> GruppoControlloUtilizzati { get; private set; }
        public static List<int> GruppoControlloCancellati { get; set; }

        public ConfiguratoreRibbon()
        {
            Utility.InitializeUtility();
            Utility.StdFont = this.Font;
            InitializeComponent();

            btnProd.Image = Iren.PSO.Base.Properties.Resources.prod;
            btnTest.Image = Iren.PSO.Base.Properties.Resources.test;

            //trovo tutte le risorse disponibili in Iren.ToolsExcel.Base
            var resourceSet = Iren.PSO.Base.Properties.Resources.ResourceManager.GetResourceSet(CultureInfo.InstalledUICulture, true, true);

            //Considero solo quelle che sono di tipo Image
            var imgs =
                from r in resourceSet.Cast<DictionaryEntry>()
                where r.Value is Image
                select r;

            foreach (var img in imgs)
            {
                Utility.ImageListNormal.Images.Add(img.Key as string, img.Value as Image);
                Utility.ImageListSmall.Images.Add(img.Key as string, img.Value as Image);
            }



            //inizializzazione connessione
            //_ambienti = Workbook.AppSettings("Ambienti").Split('|');
            btnTest.Checked = true;
            //ChangeAmbiente(btnTest, null);
            //
            //DataBase.DB.SetParameters(idUtente: 62);

        }

        private void CaricaAnteprimaRibbon()
        {
            ControlliUtilizzati = new List<int>();
            GruppiUtilizzati = new List<int>();
            GruppoControlloUtilizzati = new List<int>();
            GruppoControlloCancellati = new List<int>();

            DataTable ribbon = DataBase.Select(DataBase.SP.RIBBON.GRUPPO_CONTROLLO);
            DataTable funzioni = DataBase.Select(DataBase.SP.RIBBON.CONTROLLO_FUNZIONE);

            if (ribbon != null)
            {
                int idGroup = -1;
                RibbonGroup grp = null;
                foreach (DataRow r in ribbon.Rows)
                {
                    //prendo nota di cosa è utilizzato.
                    GruppoControlloUtilizzati.Add((int)r["IdGruppoControllo"]);
                    ControlliUtilizzati.Add((int)r["IdControllo"]);
                    GruppiUtilizzati.Add((int)r["IdGruppo"]);

                    if (!r["IdGruppo"].Equals(idGroup))
                    {
                        idGroup = (int)r["IdGruppo"];
                        grp = new RibbonGroup(panelRibbonLayout, (int)r["IdGruppo"]);
                        grp.Name = r["Nome"].ToString();
                        panelRibbonLayout.Controls.Add(grp);
                        grp.Text = r["LabelGruppo"].ToString();
                    }
                    
                    Control ctrl = Utility.AddControlToGroup(grp, r, funzioni);
                    ctrl.ContextMenuStrip = new DisabilitaMenuStrip();
                }
            }
        }

        //SPOSTAMENTI
        private void MoveDown_Click(object sender, EventArgs e)
        {
            IRibbonControl ctrl = ActiveControl as IRibbonControl;

            if (ctrl != null && ctrl.Slot < 3)
            {
                var nextCtrl = Utility.GetAll(ActiveControl.Parent)
                        .Where(c => ActiveControl.Bottom == c.Top).FirstOrDefault();

                if (nextCtrl != null)
                {
                    nextCtrl.Top = ActiveControl.Top;
                    ActiveControl.Top = nextCtrl.Bottom;
                }
            }
        }
        private void MoveUp_Click(object sender, EventArgs e)
        {
            IRibbonControl ctrl = ActiveControl as IRibbonControl;

            if (ctrl != null && ctrl.Slot < 3)
            {
                var nextCtrl = Utility.GetAll(ActiveControl.Parent)
                        .Where(c => ActiveControl.Top == c.Bottom).FirstOrDefault();

                if (nextCtrl != null)
                {
                    ActiveControl.Top = nextCtrl.Top;
                    nextCtrl.Top = ActiveControl.Bottom;
                }
            }
        }
        private void MoveLeft_Click(object sender, EventArgs e)
        {
            if(ActiveControl != null) 
            {
                var oth = ActiveControl.Parent.Controls.Cast<Control>()
                    .Where(c => c.Right == ActiveControl.Left)
                    .FirstOrDefault();

                if(oth != null)
                {
                    ActiveControl.Left = oth.Left;
                    oth.Left = ActiveControl.Right;
                }
            }
        }
        private void MoveRight_Click(object sender, EventArgs e)
        {
            if (ActiveControl != null)
            {
                var oth = ActiveControl.Parent.Controls.Cast<Control>()
                    .Where(c => c.Left == ActiveControl.Right)
                    .FirstOrDefault();

                if (oth != null)
                {
                    oth.Left = ActiveControl.Left;
                    ActiveControl.Left = oth.Right;
                }
            }
        }

        private bool IsRibbonGroupSelected()
        {
            if (ActiveControl.GetType() != typeof(RibbonGroup))
            {
                MessageBox.Show("Nessun gruppo selezionato...", "ATTENZIONE!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        private Control CreateEmptyContainer()
        {
            if (IsRibbonGroupSelected())
                return Utility.CreateEmptyContainer(ActiveControl);

            return null;   
        }

        //TASTI
        private void AggiungiNuovoTasto_Click(object sender, EventArgs e)
        {
            Control container = CreateEmptyContainer();
            if (container != null)
            {
                RibbonButton newBtn = new RibbonButton(panelRibbonLayout);
                if (newBtn.ImageKey != "")
                {
                    container.Controls.Add(newBtn);
                    newBtn.Top = container.Padding.Top;
                    newBtn.Left = container.Padding.Left;
                    ActiveControl.Controls.Add(container);
                }
            }
        }
        private void ScegliTastoEsistente_Click(object sender, EventArgs e)
        {
            if (IsRibbonGroupSelected())
            {
                using (ControlliEsistenti ctrlForm = new ControlliEsistenti(ActiveControl, 1, 2))
                {
                    ctrlForm.ShowDialog();
                }
            }
        }
        
        //COMBO
        private void AggiungiNuovoCombo_Click(object sender, EventArgs e)
        {
            Control container = CreateEmptyContainer();
            if (container != null)
            {
                RibbonDropDown newDrpDwn = new RibbonDropDown(panelRibbonLayout);
                if(!newDrpDwn.IsDisposed)
                {
                    container.Controls.Add(newDrpDwn);
                    newDrpDwn.Top = container.Padding.Top;
                    newDrpDwn.Left = container.Padding.Left;

                    ActiveControl.Controls.Add(container);
                }
            }
        }
        private void ScegliComboEsistente_Click(object sender, EventArgs e)
        {
            if (IsRibbonGroupSelected())
            {
                using (ControlliEsistenti ctrlForm = new ControlliEsistenti(ActiveControl, 3))
                {
                    ctrlForm.ShowDialog();
                }
            }
        }
        

        //GRUPPI
        private void AggiungiGruppo_Click(object sender, EventArgs e)
        {
            RibbonGroup newGroup = new RibbonGroup(panelRibbonLayout);
            if(!newGroup.IsDisposed)
                Utility.AddGroupToRibbon(panelRibbonLayout, newGroup);
        }
        private void ScegliGruppoEsistente_Click(object sender, EventArgs e)
        {
            using (GruppiEsistenti grpForm = new GruppiEsistenti(panelRibbonLayout))
            {
                grpForm.ShowDialog();
            }
        }

        //VUOTI
        private void AggiungiContenitoreVuoto_Click(object sender, EventArgs e)
        {
            Control container = CreateEmptyContainer();
            if (container != null)
                ActiveControl.Controls.Add(container);
        }

        private void ApplicaConfigurazione_Click(object sender, EventArgs e)
        {
            //rimuovo i componenti cancellati
            if(GruppoControlloCancellati.Count > 0) 
            {
                DataBase.Delete(DataBase.SP.RIBBON.DELETE_GRUPPO_CONTROLLO, "@Ids=" + string.Join(",", GruppoControlloCancellati));
            }
            //trovo tutti i gruppi
            var groups =
                panelRibbonLayout.Controls.OfType<RibbonGroup>().OrderBy(g => g.Left);
            
            int ordine = 1;
            foreach (RibbonGroup group in groups)
            {
                //trovo tutti i contenitori
                var containers =
                    group.Controls.OfType<ControlContainer>().OrderBy(c => c.Left);

                Dictionary<string, object> outP = new Dictionary<string,object>();
                int groupId = -1;
                if (DataBase.Insert(DataBase.SP.RIBBON.INSERT_GRUPPO, new Iren.PSO.Core.QryParams()
                    {
                        {"@Id", group.IdGruppo},
                        {"@Nome", group.Name},
                        {"@Label", group.Text}
                    }, out outP))
                    groupId = (int)outP["@Id"];

                foreach (var container in containers)
                {
                    //trovo tutti i controlli contenuti nei contenitori
                    var ctrls =
                        container.Controls.Cast<IRibbonControl>();

                    foreach (IRibbonControl ctrl in ctrls)
                    {
                        outP = new Dictionary<string, object>();
                        int ctrlId = -1;
                        if (DataBase.Insert(DataBase.SP.RIBBON.INSERT_CONTROLLO, new Iren.PSO.Core.QryParams()
                            {
                                {"@Id", ctrl.IdControllo},
                                {"@IdTipologiaControllo", ctrl.IdTipologia},
                                {"@Nome", ctrl.Name},
                                {"@Descrizione", ctrl.Description ?? ""},
                                {"@Immagine", ctrl.ImageKey},
                                {"@Label", ctrl.Text},
                                {"@ScreenTip", ctrl.ScreenTip ?? ""},
                                {"@ControlSize", ctrl.Dimension}
                            }, out outP))
                            ctrlId = (int)outP["@Id"];

                        if (DataBase.Insert(DataBase.SP.RIBBON.INSERT_GRUPPO_CONTROLLO, new Iren.PSO.Core.QryParams() { 
                            {"@Id", 0},
                            //{"@IdApplicazione", },
	                        //{"@IdUtente", 62},
	                        {"@IdGruppo", groupId},
	                        {"@IdControllo", ctrlId},
                            {"@Abilitato", ctrl.Enabled ? "1" : "0"},
	                        {"@Ordine", ordine++}
                        }, out outP))
                        {
                            DataBase.Delete(DataBase.SP.RIBBON.DELETE_FUNZIONI_CONTROLLO, "@IdGruppoControllo=" + outP["@Id"]);

                            int ordineFunzioni = 1;
                            foreach(int idFunzione in ctrl.Functions)
                            {
                                DataBase.Insert(DataBase.SP.RIBBON.INSERT_CONTROLLO_FUNZIONE, new Iren.PSO.Core.QryParams()
                                {
                                    {"@IdGruppoControllo", outP["@Id"]},
                                    {"@IdFunzione", idFunzione},
                                    {"@Ordine", ordineFunzioni++}
                                });
                            }
                        }
                    }
                }
            }
            //refresh dell'anteprima
            RicaricaRibbon_Click(null, null);
        }
        private void RicaricaRibbon_Click(object sender, EventArgs e)
        {
            Utility.Refreshing = true;
            panelRibbonLayout.Controls.Clear();
            CaricaAnteprimaRibbon();
            Utility.Refreshing = false;
        }

        private void CambioApplicazione(object sender, EventArgs e)
        {
            if (drpApplicazioni.SelectedValue != null)
            {
                DataBase.IdApplicazione = (int)drpApplicazioni.SelectedValue;
                if (drpUtenti.SelectedValue != null)
                    RicaricaRibbon_Click(null, null);
            }
        }

        private void CambioUtente(object sender, EventArgs e)
        {
            if (drpUtenti.SelectedValue != null)
            {
                DataBase.IdUtente = (int)drpUtenti.SelectedValue;
                if (drpApplicazioni.SelectedValue != null)
                    RicaricaRibbon_Click(null, null);
            }
        }

        private void CopiaConfigurazione_Click(object sender, EventArgs e)
        {
            using (CopiaConfigurazione form = new CopiaConfigurazione())
            {
                form.ShowDialog();
            }
        }

        private void ChangeAmbiente(object sender, EventArgs e)
        {
            ToolStripButton btn = (ToolStripButton)sender;

            if (!btn.Checked)
            {
                btn.CheckedChanged -= ChangeAmbiente;
                btn.Checked = true;
                btn.CheckedChanged += ChangeAmbiente;
                return;
            }
                
            btnProd.CheckedChanged -= ChangeAmbiente;
            btnTest.CheckedChanged -= ChangeAmbiente;

            if (btn == btnProd)
                btnTest.Checked = false;
            else
                btnProd.Checked = false;

            btnTest.CheckedChanged += ChangeAmbiente;
            btnProd.CheckedChanged += ChangeAmbiente;
            if(DataBase.IsInitialized)
                DataBase.Close();
            if (btn == btnProd)
                DataBase.CreateNew(Simboli.PROD, false);
            else if (btn == btnTest)
                DataBase.CreateNew(Simboli.TEST, false);

            drpApplicazioni.SelectedValueChanged -= CambioApplicazione;
            drpUtenti.SelectedValueChanged -= CambioUtente;
            //carico la lista di applicazioni configurabili
            DataTable applicazioni = DataBase.Select(DataBase.SP.APPLICAZIONE, "@IdApplicazione=0;@IdUtente=1");
            if (applicazioni != null)
            {
                drpApplicazioni.DisplayMember = "DesApplicazione";
                drpApplicazioni.ValueMember = "IdApplicazione";
                drpApplicazioni.DataSource = applicazioni;
            }

            //carico la lista degli utenti
            DataTable utenti = DataBase.Select(DataBase.SP.UTENTE_GRUPPO, "@IdUtenteGruppo=0");
            var allUser = utenti.AsEnumerable()
                .Where(r => r["IdUtenteGruppo"].Equals(1) || r["IdUtenteGruppo"].Equals(5))
                .Select(r => new { IdUtente = r["IdUtente"], Nome = r["Nome"] })
                .ToList();

            if (utenti != null)
            {
                drpUtenti.DisplayMember = "Nome";
                drpUtenti.ValueMember = "IdUtente";
                drpUtenti.DataSource = allUser;
            }
            DataBase.IdApplicazione = (int)drpApplicazioni.SelectedValue;
            DataBase.IdUtente = (int)drpUtenti.SelectedValue;

            drpApplicazioni.SelectedValueChanged += CambioApplicazione;
            drpUtenti.SelectedValueChanged += CambioUtente;
            
            RicaricaRibbon_Click(null, null);
        }
    }
}
