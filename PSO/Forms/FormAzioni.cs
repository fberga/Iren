using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Forms
{
    public partial class FormAzioni : Form
    {
        #region Variabili

        private DataView _azioni;
        private DataView _categorie;
        private DataView _categoriaEntita;
        private DataView _azioniCategorie;
        private DataView _entitaAzioni;
        private AEsporta _esporta;
        private ARiepilogo _r;
        private ACarica _carica;
        private List<DateTime> _toProcessDates = new List<DateTime>();
        private FormSelezioneDate selDate = new FormSelezioneDate();

        #endregion

        #region Costruttori

        public FormAzioni(AEsporta esporta, ARiepilogo riepilogo, ACarica carica)
        {
            InitializeComponent();

            _esporta = esporta;
            _r = riepilogo;
            _carica = carica;

            _categorie = Workbook.Repository[DataBase.TAB.CATEGORIA].DefaultView;
            _categorie.RowFilter = "IdApplicazione = " + Workbook.IdApplicazione;
            _categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            _categoriaEntita.RowFilter = "Gerarchia = '' OR Gerarchia IS NULL AND IdApplicazione = " + Workbook.IdApplicazione;
            _azioni = Workbook.Repository[DataBase.TAB.AZIONE].DefaultView;
            _azioni.RowFilter = "Visibile = 1 AND IdApplicazione = " + Workbook.IdApplicazione;
            _azioniCategorie = Workbook.Repository[DataBase.TAB.AZIONE_CATEGORIA].DefaultView;
            _azioniCategorie.RowFilter = "IdApplicazione = " + Workbook.IdApplicazione;
            _entitaAzioni = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE].DefaultView;
            _entitaAzioni.RowFilter =  "IdApplicazione = " + Workbook.IdApplicazione;
            DataView entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].DefaultView;
            entitaProprieta.RowFilter = "IdApplicazione = " + Workbook.IdApplicazione;

            ConfigStructure();

            if (Struct.intervalloGiorni == 0 || Struct.tipoVisualizzazione == "R")
            {
                comboGiorni.Text = Workbook.DataAttiva.ToString("dddd dd MMM yyyy");
                _toProcessDates.Add(Workbook.DataAttiva);
                comboGiorni.Enabled = false;
            }
            else
            {
                selDate.VisibleChanged += selDate_VisibleChanged;
            }

            if (!DataBase.OpenConnection())
                btnMeteo.Enabled = false;
            DataBase.CloseConnection();

            this.Text = Simboli.NomeApplicazione + " - Azioni";
            CaricaAzioni();
            CaricaCategorie();
        }

        #endregion

        #region Metodi

        private void ConfigStructure()
        {
            System.Collections.IDictionary settings = (System.Collections.IDictionary)ConfigurationManager.GetSection("formSettings/azioniForm");

            Regex falseMatch = new Regex("false|0", RegexOptions.IgnoreCase);

            if (Workbook.Repository.Applicazione["VisCategoriaRiepilogo"].Equals("0"))
            {
                panelCategorie.Hide();
                Width -= panelCategorie.Width;
            }
            if (Workbook.Repository.Applicazione["VisMeteo"].Equals("0"))
                btnMeteo.Hide();

            if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("0"))
                groupMercati.Hide();
            else
            {
                int hour = DateTime.Now.Hour;

                Dictionary<string, SpecMercato> mercati = new Dictionary<string,SpecMercato>();

                if (Workbook.IdApplicazione == 10)
                {
                    mercati = Simboli.MercatiMB;
                }
                /* 
                // inutile perchè per l'id = 18 ModificaDinamica è stata impostata a 0, e quindi viene nascosta la groupBox
                // groupMercati
                else if (Workbook.IdApplicazione == 18)
                {
                    mercati = Simboli.MercatiMI;
                }
                */

                var mercatoAttivo = mercati
                    .Select(kv => new { Nome = kv.Key, Attivo = kv.Value.Chiusura > hour });
                

                // Escludo i mercati MI1, MI2 3 MI3 in caso di irario antecedente alle 13:00 -> inutile perchè
                // per l'id = 18 ModificaDinamica è stata impostata a 0, e quindi viene nascosta la groupBox groupMercati 
                /*if (Workbook.IdApplicazione == 18)
                {
                    if (hour < Simboli.MercatiMI["MI7"].Chiusura)
                    {
                        mercatoAttivo = mercati.Select(kv => new { Nome = kv.Key, Attivo = kv.Value.Chiusura > hour && !(kv.Key == "MI1" || kv.Key == "MI2" || kv.Key == "MI3") });                        
                    }
                }*/

                int leftInc = groupMercati.Width / mercati.Count;

                int i = 0;
                foreach (var mk in mercatoAttivo)
                {
                    CheckBox chk = new CheckBox()
                    {
                        Name = "check" + mk.Nome,
                        Text = mk.Nome,
                        AutoSize = true,
                        Enabled = mk.Attivo,
                        Checked = mk.Attivo,
                        Location = new Point(3 + leftInc * i++, (groupMercati.Height / 2) - 5)
                    };
                    groupMercati.Controls.Add(chk);
                }
            }
        }
        private void CaricaAzioni()
        {
            var stato = DataBase.StatoDB;

            foreach (DataRowView azione in _azioni)
            {
                bool aggiungi = true;
                if (azione["IdFonte"] != DBNull.Value)
                {
                    var fonte = (Core.DataBase.NomiDB)Enum.Parse(typeof(Core.DataBase.NomiDB), azione["IdFonte"].ToString());
                    aggiungi = stato[fonte] == ConnectionState.Open;
                }

                if (azione["Operativa"].Equals("0") || (azione["Gerarchia"] is DBNull && aggiungi))
                {
                    treeViewAzioni.Nodes.Add(azione["SiglaAzione"].ToString(), azione["DesAzione"].ToString());
                }
                else if (aggiungi)
                {
                    if(treeViewAzioni.Nodes.ContainsKey(azione["Gerarchia"].ToString()))
                        treeViewAzioni.Nodes[azione["Gerarchia"].ToString()].Nodes.Add(azione["SiglaAzione"].ToString(), azione["DesAzione"].ToString());
                }
            }

            int i = 0;
            while (i < treeViewAzioni.Nodes.Count)
            {
                if (treeViewAzioni.Nodes[i].Nodes.Count == 0)
                    treeViewAzioni.Nodes.RemoveAt(i);
                else
                    i++;
            }

            treeViewAzioni.ExpandAll();
        }
        private void CaricaCategorie()
        {
            foreach (DataRowView categoria in _categorie)
            {
                if (categoria["Operativa"].Equals("0"))
                {
                    treeViewCategorie.Nodes.Add(categoria["SiglaCategoria"].ToString(), categoria["DesCategoriaBreve"].ToString());
                }
                else
                {
                    if (categoria["Gerarchia"] is DBNull)
                        treeViewCategorie.Nodes.Add(categoria["SiglaCategoria"].ToString(), categoria["DesCategoriaBreve"].ToString());
                    else
                        treeViewCategorie.Nodes[categoria["Gerarchia"].ToString()].Nodes.Add(categoria["SiglaCategoria"].ToString(), categoria["DesCategoriaBreve"].ToString());
                }
            }
            treeViewCategorie.ExpandAll();
        }
        private void ThroughAllNodes(TreeNodeCollection root, Action<TreeNode> callback)
        {
            if (root.Count > 0)
            {
                foreach (TreeNode node in root.OfType<TreeNode>())
                {
                    callback(node);
                    ThroughAllNodes(node.Nodes, callback);
                }
            }
        }
        private void CaricaEntita()
        {
            Dictionary<string, bool> notSel = new Dictionary<string, bool>();

            foreach (TreeNode node in treeViewUP.Nodes)
            {
                if (!node.Checked)
                    notSel.Add(node.Name, false);
            }

            treeViewUP.Nodes.Clear();

            ThroughAllNodes(treeViewCategorie.Nodes, n =>
            {
                if (n.Checked)
                {
                    _categoriaEntita.RowFilter = "SiglaCategoria = '" + n.Name + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                    _categoriaEntita.Sort = "DesEntita";
                    foreach (DataRowView entita in _categoriaEntita)
                    {
                        ThroughAllNodes(treeViewAzioni.Nodes, n1 =>
                        {
                            if (n1.Checked)
                            {
                                _entitaAzioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaAzione = '" + n1.Name + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                                if (_entitaAzioni.Count > 0 && treeViewUP.Nodes.Find(entita["SiglaEntita"].ToString(), true).Length == 0)
                                {
                                    treeViewUP.Nodes.Add(entita["SiglaEntita"].ToString(), entita["DesEntita"].ToString());
                                    if (notSel.ContainsKey(entita["SiglaEntita"].ToString()))
                                        treeViewUP.Nodes[entita["SiglaEntita"].ToString()].Checked = false;
                                    else
                                        treeViewUP.Nodes[entita["SiglaEntita"].ToString()].Checked = true;
                                }
                            }
                        });
                    }
                    _categoriaEntita.Sort = "";
                }
            });
        }
        private void Evidenzia(TreeNode node, bool evidenzia)
        {
            if (evidenzia)
            {
                node.BackColor = System.Drawing.Color.Gold;
                node.ForeColor = System.Drawing.Color.DarkRed;
            }
            else
            {
                node.BackColor = treeViewAzioni.BackColor;
                node.ForeColor = treeViewAzioni.ForeColor;
                node.NodeFont = treeViewAzioni.Font;
            }
        }
        private void ColoraNodi()
        {
            ThroughAllNodes(treeViewAzioni.Nodes, n =>
            {
                Evidenzia(n, n.Checked);
            });
            ThroughAllNodes(treeViewCategorie.Nodes, n =>
            {
                Evidenzia(n, n.Checked);
            });
            ThroughAllNodes(treeViewUP.Nodes, n =>
            {
                Evidenzia(n, n.Checked);
            });
        }
        private void CheckParents()
        {
            foreach (TreeNode node in treeViewAzioni.Nodes)
            {
                if (node.Nodes.Count != 0)
                {
                    if (HasCheckedNode(node))
                        node.Checked = true;
                    else
                        node.Checked = false;
                }
            }
            foreach (TreeNode node in treeViewCategorie.Nodes)
            {
                if (node.Nodes.Count != 0)
                {
                    if (HasCheckedNode(node))
                        node.Checked = true;
                    else
                        node.Checked = false;
                }
            }
        }
        private bool HasCheckedNode(TreeNode node)
        {
            return node.Nodes.Cast<TreeNode>().Any(n => n.Checked);
        }

        public DialogResult ShowDialog(bool fromConsole, string[] azioni, string[] entita)
        {
            if (fromConsole)
            {
                if(azioni != null && azioni.Length > 0)
                {
                    foreach (string a in azioni)
                        treeViewAzioni.Nodes.Find(a, true)[0].Checked = true; ;
                }
                else
                {
                    foreach (TreeNode n in treeViewAzioni.Nodes)
                        n.Checked = true;
                }
                if(entita != null && entita.Length > 0)
                {
                    foreach (TreeNode n in treeViewUP.Nodes)
                        n.Checked = false;
                    
                    foreach (string e in entita)
                        treeViewUP.Nodes.Find(e, true)[0].Checked = true; ;
                }

                selDate.SelectAll();
                btnApplica_Click(null, null);
                return System.Windows.Forms.DialogResult.OK;
            }

            return base.ShowDialog();
        }

        #endregion

        #region Eventi

        private void frmAZIONI_Load(object sender, EventArgs e)
        {
            
        }
        private void treeView_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            e.Cancel = true;
        }
        private void treeView_AfterCheck(object sender, TreeViewEventArgs e)
        {
            TreeView from = sender as TreeView;
            TreeView to = from.Name == "treeViewAzioni" ? treeViewCategorie : treeViewAzioni;

            from.AfterCheck -= treeView_AfterCheck;
            to.AfterCheck -= treeView_AfterCheck;

            List<TreeNode> justChecked = new List<TreeNode>();
            if (e.Node.Nodes.Count > 0)
                foreach (TreeNode node in e.Node.Nodes)
                {
                    node.Checked = e.Node.Checked;
                    justChecked.Add(node);
                }
            else
                justChecked.Add(e.Node);

            string filter = from.Name == "treeViewAzioni" ? "SiglaAzione" : "SiglaCategoria";
            string field = from.Name == "treeViewAzioni" ? "SiglaCategoria" : "SiglaAzione";

            
            if (e.Node.Checked)
            {
                foreach (TreeNode node in justChecked)
                {
                    _azioniCategorie.RowFilter = filter + " = '" + node.Name + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                    foreach (DataRowView azioneCategoria in _azioniCategorie)
                        to.Nodes.Find(azioneCategoria[field].ToString(), true)[0].Checked = true;
                }
            }
            else
            {
                Dictionary<string, bool> checkedNodes = new Dictionary<string, bool>();
                ThroughAllNodes(from.Nodes, n =>
                {
                    if (n.Nodes.Count == 0)
                    {
                        _azioniCategorie.RowFilter = filter + " = '" + n.Name + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                        foreach (DataRowView azioneCategoria in _azioniCategorie)
                        {
                            if (checkedNodes.ContainsKey(azioneCategoria[field].ToString()))
                                checkedNodes[azioneCategoria[field].ToString()] = checkedNodes[azioneCategoria[field].ToString()] || n.Checked;
                            else
                                checkedNodes.Add(azioneCategoria[field].ToString(), n.Checked);
                        }
                    }
                });

                ThroughAllNodes(to.Nodes, n =>
                {
                    if (n.Nodes.Count == 0)
                        n.Checked = n.Checked && checkedNodes[n.Name];
                });
            }

            CheckParents();
            CaricaEntita();
            ColoraNodi();

            to.AfterCheck += treeView_AfterCheck;
            from.AfterCheck += treeView_AfterCheck;
        }
        private void treeViewUP_AfterCheck(object sender, TreeViewEventArgs e)
        {
            checkTutte.CheckedChanged -= checkTutte_CheckedChanged;
            
            Evidenzia(e.Node, e.Node.Checked);

            bool check = true;
            foreach (TreeNode node in treeViewUP.Nodes)
            {
                check = check && node.Checked;
            }
            checkTutte.Checked = check;

            checkTutte.CheckedChanged += checkTutte_CheckedChanged;
        }
        private void btnMeteo_Click(object sender, EventArgs e)
        {
            if (_toProcessDates.Count > 0)
            {
                if ((_toProcessDates.Count > 1 && MessageBox.Show("Ci sono più date selezionate. Procedere con la prima?", Simboli.NomeApplicazione, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes) || _toProcessDates.Count == 1)
                {
                    FormMeteo meteo = new FormMeteo(_toProcessDates[0], _carica, _r);
                    meteo.ShowDialog();
                }
            }
            else
                MessageBox.Show("Non è stata selezionata alcuna data...", Simboli.NomeApplicazione, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        private void btnApplica_Click(object sender, EventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            btnApplica.Enabled = false;
            btnAnnulla.Enabled = false;
            btnMeteo.Enabled = false;

            if (_toProcessDates.Count == 0)
                MessageBox.Show("Non è stata selezionata alcuna data...", Simboli.NomeApplicazione, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else if (treeViewUP.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length == 0)
                MessageBox.Show("Non è stata selezionata alcuna unità...", Simboli.NomeApplicazione, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                SplashScreen.Show();
                Workbook.Application.EnableEvents = false;
                Workbook.ScreenUpdating = false;
                Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

                bool caricaOrGenera = false;
                bool loaded = false;

                bool g_mp_mgp = false;

                DataView entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].DefaultView;

                ThroughAllNodes(treeViewAzioni.Nodes, nodoAzione =>
                {
                    if (loaded && Regex.Match(nodoAzione.Name, @"\w[^\d]+").Value == "GENERA")
                    {
                        Workbook.Application.CalculateFull();
                        Workbook.ScreenUpdating = false;
                        loaded = false;
                    }

                    if (nodoAzione.Checked && nodoAzione.Nodes.Count == 0)
                    {
                        TreeNode[] nodiEntita = treeViewUP.Nodes.OfType<TreeNode>().Where(node => node.Checked).ToArray();
                        _azioni.RowFilter = "SiglaAzione = '" + nodoAzione.Name + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                        if (!g_mp_mgp)
                            g_mp_mgp = nodoAzione.Name == "G_MP_MGP";

                        ThroughAllNodes(treeViewUP.Nodes, nodoEntita =>
                        {
                            foreach (DateTime date in _toProcessDates)
                            {
                                string suffissoData = Date.GetSuffissoData(date);

                                if (nodoEntita.Checked && nodoEntita.Nodes.Count == 0)
                                {
                                    entitaProprieta.RowFilter = "SiglaEntita = '" + nodoEntita.Name + "' AND SiglaProprieta LIKE '%GIORNI_STRUTTURA' AND IdApplicazione = " + Workbook.IdApplicazione;
                                    int intervalloGiorni = Struct.intervalloGiorni;
                                    if (entitaProprieta.Count > 0)
                                        intervalloGiorni = int.Parse("" + entitaProprieta[0]["Valore"]);

                                    if (date <= Workbook.DataAttiva.AddDays(intervalloGiorni))
                                    {
                                        string nomeFoglio = DefinedNames.GetSheetName(nodoEntita.Name);
                                        bool presente;

                                        DataView entitaAzione = new DataView(Workbook.Repository[DataBase.TAB.ENTITA_AZIONE]);
                                        entitaAzione.RowFilter = "SiglaEntita = '" + nodoEntita.Name + "' AND SiglaAzione = '" + nodoAzione.Name + "' AND IdApplicazione = " + Workbook.IdApplicazione;

                                        if (entitaAzione.Count > 0 && (entitaAzione[0]["Giorno"] is DBNull || entitaAzione[0]["Giorno"].ToString().Contains(Date.GetSuffissoData(date))))
                                        {
                                            SplashScreen.UpdateStatus("[" + date.ToShortDateString() + "] " + nodoAzione.Parent.Text + " " + nodoAzione.Text + ": " + nodoEntita.Text);

                                            //string[] mercati = null;
                                            string[] mercati = null;
                                            if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1"))
                                            {
                                                mercati = groupMercati.Controls
                                                    .OfType<CheckBox>()
                                                    .Where(c => c.Checked)
                                                    .Select(c => Regex.Match(c.Text, @"\d+").Value)
                                                    .OrderBy(s => s)
                                                    .ToArray();
                                            }
                                            else if (Workbook.IdApplicazione == 18)
                                            {
                                                mercati = new string[] { Regex.Match(Workbook.Mercato, @"\d+").Value };
                                            }

                                            switch (Regex.Match(nodoAzione.Parent.Name, @"\w[^\d]+").Value)
                                            {
                                                case "CARICA":
                                                    presente = _carica.AzioneInformazione(nodoEntita.Name, nodoAzione.Name, nodoAzione.Parent.Name, date, mercati);
                                                    _r.AggiornaRiepilogo(nodoEntita.Name, nodoAzione.Name, presente, date);
                                                    caricaOrGenera = true;
                                                    loaded = true;
                                                    break;
                                                case "GENERA":
                                                    presente = _carica.AzioneInformazione(nodoEntita.Name, nodoAzione.Name, nodoAzione.Parent.Name, date, mercati);
                                                    _r.AggiornaRiepilogo(nodoEntita.Name, nodoAzione.Name, presente, date);
                                                    caricaOrGenera = true;
                                                    break;
                                                case "ESPORTA":
                                                    presente = _esporta.RunExport(nodoEntita.Name, nodoAzione.Name, nodoEntita.Text, nodoAzione.Text, date, mercati);
                                                    if (presente)
                                                        _r.AggiornaRiepilogo(nodoEntita.Name, nodoAzione.Name, presente, date);
                                                    break;
                                            }

                                            if (_azioni[0]["Relazione"] != DBNull.Value && Struct.visualizzaRiepilogo)
                                            {
                                                string[] azioneRelazione = _azioni[0]["Relazione"].ToString().Split(';');

                                                DefinedNames definedNames = new DefinedNames("Main");

                                                foreach (string relazione in azioneRelazione)
                                                {
                                                    _azioni.RowFilter = "SiglaAzione = '" + relazione + "' AND IdApplicazione = " + Workbook.IdApplicazione;

                                                    Range rng = new Range(definedNames.GetRowByName(nodoEntita.Name), definedNames.GetColFromName(relazione, suffissoData));
                                                    if (Workbook.Main.Range[rng.ToString()].Interior.ColorIndex != 2)
                                                    {
                                                        Workbook.Main.Range[rng.ToString()].Value = "RI" + _azioni[0]["Gerarchia"];
                                                        Style.RangeStyle(Workbook.Main.Range[rng.ToString()], fontSize: 8, bold: true, foreColor: 3, backColor: 6, align: Excel.XlHAlign.xlHAlignCenter);
                                                    }
                                                }
                                                _azioni.RowFilter = "SiglaAzione = '" + nodoAzione.Name + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                                            }
                                        }
                                    }
                                }
                            }
                        });
                        switch (Regex.Match(nodoAzione.Parent.Name, @"\w[^\d]+").Value)
                        {
                            case "CARICA":
                                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogCarica, "Carica: " + nodoAzione.Text);
                                break;
                            case "GENERA":
                                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogGenera, "Genera: " + nodoAzione.Text);
                                break;
                            case "ESPORTA":
                                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogEsporta, "Esporta: " + nodoAzione.Text);
                                break;
                        }
                    }
                });

                Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                if (caricaOrGenera)
                {
                    SplashScreen.UpdateStatus("Salvo su DB");
                    Sheet.SalvaModifiche();
                    DataBase.SalvaModificheDB();
                    if (Workbook.IdApplicazione == 5 && g_mp_mgp)
                    {
                        SplashScreen.UpdateStatus("Aggiorno rendimenti");
                        Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
                        
                        ThroughAllNodes(treeViewUP.Nodes, nodoEntita =>
                        {
                            if (nodoEntita.Checked && nodoEntita.Nodes.Count == 0)
                            {
                                _categoriaEntita.RowFilter = "SiglaEntita = '" + nodoEntita.Name + "' AND IdApplicazione = " + Workbook.IdApplicazione;

                                if (_categoriaEntita[0]["SiglaCategoria"].Equals("IREN_60T"))
                                {
                                    foreach (DateTime date in _toProcessDates)
                                    {
                                        _carica.AzioneInformazione(nodoEntita.Name, "RENDIMENTO", "CARICA", date, null);
                                    }
                                }
                            }
                        });
                        Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                        SplashScreen.UpdateStatus("Salvo su DB");
                        Sheet.SalvaModifiche();
                        DataBase.SalvaModificheDB();
                    }
                    

                }

                Workbook.Application.EnableEvents = true;
                Workbook.ScreenUpdating = true;
                SplashScreen.Close();
            }

            btnApplica.Enabled = true;
            btnAnnulla.Enabled = true;
            btnMeteo.Enabled = true;

            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        private void checkTutte_CheckedChanged(object sender, EventArgs e)
        {
            treeViewUP.AfterCheck -= treeViewUP_AfterCheck;
            foreach (TreeNode node in treeViewUP.Nodes)
            {
                node.Checked = checkTutte.Checked;
                Evidenzia(node, checkTutte.Checked);
            }
            treeViewUP.AfterCheck += treeViewUP_AfterCheck;
        }
        private void FormAzioni_FormClosing(object sender, FormClosingEventArgs e)
        {
            selDate.Close();
        }
        private void comboGiorni_MouseClick(object sender, EventArgs e)
        {
            selDate.Top = comboGiorni.PointToScreen(Point.Empty).Y + comboGiorni.Height;
            selDate.Left = comboGiorni.PointToScreen(Point.Empty).X;
            selDate.Width = comboGiorni.Width;
            selDate.ShowDialog(this);
        }
        private void comboGiorni_TextChanged(object sender, EventArgs e)
        {
            comboGiorni.TextChanged -= comboGiorni_TextChanged;
            if (comboGiorni.Text == "" || comboGiorni.Text == "- Click per selezionare le date -")
            {
                comboGiorni.Text = "- Click per selezionare le date -";
                comboGiorni.Font = new Font(comboGiorni.Font, FontStyle.Italic);
                comboGiorni.ForeColor = SystemColors.ControlDarkDark;
            }
            else
            {
                comboGiorni.Font = new Font(comboGiorni.Font, FontStyle.Regular);
                comboGiorni.ForeColor = SystemColors.ControlText;
            }
            comboGiorni.TextChanged += comboGiorni_TextChanged;
        }
        private void selDate_VisibleChanged(object sender, EventArgs e)
        {
            if (!_toProcessDates.SequenceEqual(selDate.SelectedDates))
            {
                _toProcessDates = new List<DateTime>(selDate.SelectedDates);

                comboGiorni.Text = "";
                if (_toProcessDates.Count == 1)
                    comboGiorni.Text = _toProcessDates[0].ToString("ddd dd MMM yyyy");
                else if (_toProcessDates.Count > 0)
                    comboGiorni.Text = _toProcessDates.Count + " giorni";
                else
                    comboGiorni.Text = "";
            }
        }

        #endregion
    }
}
