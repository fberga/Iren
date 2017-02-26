using Iren.PSO.Base;
using Iren.PSO.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualStudio.Tools.Applications;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
//19/01/2017 FIX DomainUloadedException
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

// ***************************************************** RIBBON ***************************************************** //

namespace Iren.PSO.Applicazioni
{
    public partial class ToolsExcelRibbon
    {
        //19/01/2017 FIX DomainUloadedException
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        //19/01/2017 FIX DomainUloadedException
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern Boolean PostMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);


        #region Variabili
        
        /// <summary>
        /// Indica se tutti i tasti (a parte Aggiorna Struttura) sono disabilitati.
        /// </summary>
        private bool _allDisabled = false;
        /// <summary>
        /// Componente da aggiungere all'actionsPane del documento.
        /// </summary>
        private Forms.ErrorPane _errorPane = new Forms.ErrorPane();
        /// <summary>
        /// Variabile per svolgere delle azioni custom coi ceck.
        /// </summary>
        private Check _checkFunctions = new Check();
        /// <summary>
        /// Classe per l'aggiunta di azioni custom dopo la modifica di un Range.
        /// </summary>
        public Modifica _modificaCustom = new Modifica();

        public Forms.FormIncremento _frmInc;
        public Forms.FormIncrementoMI _frmIncMI;
        public Forms.FormBilanciamento _frmBalance;

        //private int _idApplicazione = -1;
        //private int _idUtente = -1;
        //private string _nomeUtente = "";
        private bool _updated = false;

        private DataTable _dtControllo = null;
        private DataTable _dtControlloApplicazione = null;
        private DataTable _dtFunzioni = null;

        #endregion

        #region Proprietà

        /// <summary>
        /// Proprietà che permette l'indicizzazione per nome dei vari tasti della barra Ribbon. 
        /// La necessità di questa proprietà deriva dalla necessità di abilitare/disabilitare/nascondere i tasti leggendo i parametri del DB.
        /// </summary>
        public ControlCollection Controls { get; private set; }
        public GroupsCollection Groups { get; private set; }

        #endregion

        #region Initialize 2

        public void InitializeComponent2()
        {
#if DEBUG
            //string path = @"D:\Repository\Iren\PSO\Applicazioni\" + System.AppDomain.CurrentDomain.FriendlyName.Split('.')[0] + @"\bin\Debug\" + System.AppDomain.CurrentDomain.FriendlyName;
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, System.AppDomain.CurrentDomain.FriendlyName);
            //System.Windows.Forms.MessageBox.Show(path);
#else
            string path = Environment.ExpandEnvironmentVariables(Path.Combine(Simboli.LocalBasePath, System.AppDomain.CurrentDomain.FriendlyName));
#endif
            string tmpCopy = Environment.ExpandEnvironmentVariables(@"%TEMP%\tmpRibbonLayout_" + DateTime.Now.Ticks + ".xlsm");

            File.Copy(path, tmpCopy, true);

            DataSet dsRibbonLayout = null;
            int idApplicazione = -1;
            int idUtente = -1;
#if DEBUG
            string ambiente = Simboli.TEST;
#else
            //TODO passare a PROD quando rilasciata versione ufficiale!!!
            //string ambiente = Simboli.PROD;
            string ambiente = Simboli.TEST;
#endif
            
            using (ServerDocument xls = new ServerDocument(tmpCopy))
            {
                CachedDataHostItem dataHostItem1 =
                    xls.CachedData.HostItems["Iren.PSO.Applicazioni.ThisWorkbook"];

                CachedDataItem cachedAmbiente = dataHostItem1.CachedData["ambiente"];
                if (cachedAmbiente.Xml != null)
                {
                    using (System.IO.StringReader stringReader = new System.IO.StringReader(cachedAmbiente.Xml))
                    {
                        try
                        {
                            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(string));
                            ambiente = (string)serializer.Deserialize(stringReader);
                        }
                        catch { }
                        
                    }
                }
                //inizializzo connessione con parametri temporanei
                DataBase.CreateNew(ambiente);

                CachedDataItem cachedIdApplicazione = dataHostItem1.CachedData["idApplicazione"];
                if (cachedIdApplicazione.Xml == null)
                {
                    idApplicazione = int.Parse(Workbook.AppSettings("AppID"));
                }
                else
                {
                    using (System.IO.StringReader stringReader = new System.IO.StringReader(cachedIdApplicazione.Xml))
                    {
                        try
                        {
                            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(int));
                            idApplicazione = (int)serializer.Deserialize(stringReader);
                        }
                        catch { }
                    }
                }

                CachedDataItem cachedIdUtente = dataHostItem1.CachedData["idUtente"];
                string nomeUtente = "";
                if (cachedIdUtente.Xml == null)
                {
                    Workbook.GetUtente(out idUtente, out nomeUtente);
                }
                else
                {
                    using (System.IO.StringReader stringReader = new System.IO.StringReader(cachedIdUtente.Xml))
                    {
                        try
                        {
                            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(int));
                            idUtente = (int)serializer.Deserialize(stringReader);
                        }
                        catch { }
                    }
                }

                CachedDataItem ribbonLayout = dataHostItem1.CachedData["ribbonDataSet"];
                if (ribbonLayout.Schema != null && ribbonLayout.Xml != null)
                {
                    using (System.IO.StringReader schemaReader = new System.IO.StringReader(ribbonLayout.Schema))
                    {
                        using (System.IO.StringReader xmlReader = new System.IO.StringReader(ribbonLayout.Xml))
                        {
                            dsRibbonLayout = new DataSet();
                            dsRibbonLayout.ReadXmlSchema(schemaReader);
                            dsRibbonLayout.ReadXml(xmlReader);
                        }
                    }
                }
            }            

            File.Delete(tmpCopy);

            if (dsRibbonLayout != null)
            {
                _dtControllo = dsRibbonLayout.Tables[DataBase.TAB.RIBBON.GRUPPO_CONTROLLO];
                _dtControlloApplicazione = dsRibbonLayout.Tables[DataBase.TAB.RIBBON.CONTROLLO_APPLICAZIONE];
                _dtFunzioni = dsRibbonLayout.Tables[DataBase.TAB.RIBBON.CONTROLLO_FUNZIONE];
            }
            
            if(DataBase.OpenConnection())
            {
                _updated = true;
                _dtControllo = DataBase.Select(DataBase.SP.RIBBON.GRUPPO_CONTROLLO, "@IdApplicazione=" + idApplicazione + ";@IdUtente=" + idUtente);
                _dtControllo.TableName = DataBase.TAB.RIBBON.GRUPPO_CONTROLLO;

                _dtControlloApplicazione = DataBase.Select(DataBase.SP.RIBBON.CONTROLLO_APPLICAZIONE);
                _dtControlloApplicazione.TableName = DataBase.TAB.RIBBON.CONTROLLO_APPLICAZIONE;

                _dtFunzioni = DataBase.Select(DataBase.SP.RIBBON.CONTROLLO_FUNZIONE);
                _dtFunzioni.TableName = DataBase.TAB.RIBBON.CONTROLLO_FUNZIONE;
            }

            Microsoft.Office.Tools.Ribbon.RibbonGroup grp = this.Factory.CreateRibbonGroup();

            Groups = new GroupsCollection(this);
            Controls = new ControlCollection(this);

            int idGruppo = -1;

            foreach (DataRow r in _dtControllo.Rows)
            {
                if (!r["IdGruppo"].Equals(idGruppo))
                {
                    idGruppo = (int)r["IdGruppo"];
                    grp = this.Factory.CreateRibbonGroup();
                    grp.Name = r["NomeGruppo"].ToString();
                    grp.Label = r["LabelGruppo"].ToString();

                    FrontOffice.Groups.Add(grp);
                    Groups.Add(grp);
                }

                RibbonControl ctrl = null;

                if (typeof(RibbonButton).FullName.Equals(r["SiglaTipologiaControllo"]))
                {
                    RibbonButton newBtn = this.Factory.CreateRibbonButton();

                    newBtn.ControlSize = (Microsoft.Office.Core.RibbonControlSize)r["ControlSize"];
                    newBtn.Image = (System.Drawing.Image)PSO.Base.Properties.Resources.ResourceManager.GetObject(r["Immagine"].ToString());
                    newBtn.Label = r["Label"].ToString();
                    newBtn.Name = r["Nome"].ToString();
                    newBtn.Description = r["Descrizione"].ToString();
                    newBtn.ScreenTip = r["ScreenTip"].ToString();
                    newBtn.ShowImage = true;
                    grp.Items.Add(newBtn);
                    ctrl = newBtn;
                }
                else if (typeof(RibbonToggleButton).FullName.Equals(r["SiglaTipologiaControllo"]))
                {
                    RibbonToggleButton newTglBtn = this.Factory.CreateRibbonToggleButton();

                    newTglBtn.ControlSize = (Microsoft.Office.Core.RibbonControlSize)r["ControlSize"];
                    newTglBtn.Image = (System.Drawing.Image)PSO.Base.Properties.Resources.ResourceManager.GetObject(r["Immagine"].ToString());
                    newTglBtn.Label = r["Label"].ToString();
                    newTglBtn.Name = r["Nome"].ToString();
                    newTglBtn.Description = r["Descrizione"].ToString();
                    newTglBtn.ScreenTip = r["ScreenTip"].ToString();
                    newTglBtn.ShowImage = true;

                    var ctrlIdApplicazione = _dtControlloApplicazione.AsEnumerable()
                        .Where(ctrlApp => ctrlApp["IdControllo"].Equals(r["IdControllo"]))
                        .Select(ctrlApp => (int)ctrlApp["IdApplicazione"])
                        .FirstOrDefault();
                    newTglBtn.Tag = ctrlIdApplicazione;

                    newTglBtn.Checked = ctrlIdApplicazione == idApplicazione;

                    grp.Items.Add(newTglBtn);
                    ctrl = newTglBtn;
                }
                else if (typeof(RibbonDropDown).FullName.Equals(r["SiglaTipologiaControllo"]))
                {
                    RibbonLabel lb = this.Factory.CreateRibbonLabel();
                    lb.Label = r["Label"].ToString();
                    RibbonDropDown cmb = this.Factory.CreateRibbonDropDown();
                    cmb.ShowLabel = false;
                    //cmb.Text = null;
                    cmb.Name = r["Nome"].ToString();

                    grp.Items.Add(lb);
                    grp.Items.Add(cmb);
                    ctrl = cmb;
                }
                ctrl.Enabled = r["Abilitato"].Equals("1");

                //aggiungo l'evento al controllo appena creato
                var funzioni =
                    from funz in _dtFunzioni.AsEnumerable()
                    where funz["IdGruppoControllo"].Equals(r["IdGruppoControllo"])
                    select funz;

                foreach (DataRow f in funzioni)
                {
                    try
                    {
                        EventInfo ei = ctrl.GetType().GetEvent(f["Evento"].ToString());
                        MethodInfo hi = GetType().GetMethod(f["NomeFunzione"].ToString(), BindingFlags.Instance | BindingFlags.NonPublic);
                        Delegate d = Delegate.CreateDelegate(ei.EventHandlerType, this, hi);
                        ei.AddEventHandler(ctrl, d);
                    }
                    catch (System.ArgumentException)
                    {
                        System.Windows.Forms.MessageBox.Show("Una delle funzioni collegate ai tasti non è definita!", Simboli.NomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    }
                }
                Controls.Add(ctrl);
            }

#if !DEBUG
            this.TabHome.Visible = false;
            this.TabInsert.Visible = false;
            this.TabPageLayoutExcel.Visible = false;
            this.TabFormulas.Visible = false;
            this.TabData.Visible = false;
            this.TabReview.Visible = false;
            this.TabView.Visible = false;
            this.TabDeveloper.Visible = false;
            this.TabAddIns.Visible = false;
            this.TabPrintPreview.Visible = false;
            this.TabBackgroundRemoval.Visible = false;
            this.TabSmartArtToolsDesign.Visible = false;
#endif
        
        }

        #endregion



        private void FillcmbMSD()
        {
            RibbonDropDownItem selItem = null;
            ((RibbonDropDown)Controls["cmbMSD"]).Items.Clear();
            foreach (DataRow mercato in Workbook.Repository[DataBase.TAB.MERCATI].Rows)
            {
                RibbonDropDownItem i = Factory.CreateRibbonDropDownItem();
                i.Label = mercato["DesMercato"].ToString();
                i.Tag = mercato["IdApplicazioneMercato"];
                ((RibbonDropDown)Controls["cmbMSD"]).Items.Add(i);
                if(mercato["DesMercato"].Equals(Workbook.Mercato))
                    selItem = i;
            }

            ((RibbonDropDown)Controls["cmbMSD"]).SelectionChanged -= cmbMSD_SelectionChanged;
            ((RibbonDropDown)Controls["cmbMSD"]).SelectedItem = selItem;
            ((RibbonDropDown)Controls["cmbMSD"]).SelectionChanged += cmbMSD_SelectionChanged;
        }

        private void FillcmbStagioni()
        {
            RibbonDropDownItem selItem = null;
            foreach (DataRow stagione in Workbook.Repository[DataBase.TAB.STAGIONE].Rows)
            {
                RibbonDropDownItem i = Factory.CreateRibbonDropDownItem();
                i.Label = stagione["DesTipologiaStagione"].ToString();
                i.Tag = stagione["IdTipologiaStagione"];
                ((RibbonDropDown)Controls["cmbStagione"]).Items.Add(i);
                if (stagione["DesTipologiaStagione"].Equals(Workbook.Stagione))
                    selItem = i;
            }

            ((RibbonDropDown)Controls["cmbStagione"]).SelectionChanged -= cmbStagione_SelectionChanged;
            ((RibbonDropDown)Controls["cmbStagione"]).SelectedItem = selItem;
            ((RibbonDropDown)Controls["cmbStagione"]).SelectionChanged += cmbStagione_SelectionChanged;
        }

        //09/02/2017 MOD: aggiunto combo mercati MI
        private void FillcmbMI()
        {
            RibbonDropDownItem selItem = null;
            ((RibbonDropDown)Controls["MercatoMI"]).Items.Clear();
            /*
            foreach (string mercato in Simboli.MercatiMI.Keys)
            */
            int hour = DateTime.Now.Hour;
            foreach (string mercato in Simboli.GetActiveMarkets(hour))
            {
                RibbonDropDownItem i = Factory.CreateRibbonDropDownItem();
                i.Label = mercato;
                ((RibbonDropDown)Controls["MercatoMI"]).Items.Add(i);
                if (mercato.Equals(Workbook.Mercato))
                    selItem = i;
            }

            //((RibbonDropDown)Controls["MercatoMI"]).SelectionChanged -= cmbMI_SelectionChanged;
            ((RibbonDropDown)Controls["MercatoMI"]).SelectedItem = selItem;
            //((RibbonDropDown)Controls["MercatoMI"]).SelectionChanged += cmbMI_SelectionChanged;
        }



        #region Eventi

        /// <summary>
        /// Al caricamento del Ribbon imposta i tasti e la tab da visualizzare
        /// </summary>       
        private void ToolsExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //19/01/2017 FIX DomainUloadedException
            if (Workbook.DaConsole)
            {
                //cerca la finestra ed invia il segnale di chiusura
                int WM_CLOSE = 16;
                IntPtr HWnd = FindWindow(null, "Microsoft Excel - " + Globals.ThisWorkbook.Application.ActiveWindow.Caption);
                PostMessage(HWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                //avvia la chiusura del file
                Globals.ThisWorkbook.Close(true, Type.Missing, Type.Missing);
            }
        }

        private void StatoDB_Changed(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (DataBase.OpenConnection())
            {
                //Disabilito tasti di importazione ed esportazione che funzionano solo in emergenza
                if (Controls.Contains("btnEsportaXML"))
                    Controls["btnEsportaXML"].Enabled = false;
                if (Controls.Contains("btnImportaXML"))
                    Controls["btnImportaXML"].Enabled = false;

                //Se da default i controlli sono abilitati, li abilito
                if (Controls.IsDefaultEnabled("btnProduzione"))
                    Controls["btnProduzione"].Enabled = true;
                if (Controls.IsDefaultEnabled("btnTest"))
                    Controls["btnTest"].Enabled = true;
                if (Controls.IsDefaultEnabled("btnDev"))
                    Controls["btnDev"].Enabled = true;
                if (Controls.IsDefaultEnabled("btnAggiornaDati"))
                    Controls["btnAggiornaDati"].Enabled = true;
                if (Controls.IsDefaultEnabled("btnAggiornaStruttura"))
                    Controls["btnAggiornaStruttura"].Enabled = true;
                if (Controls.IsDefaultEnabled("btnConfiguraParametri"))
                    Controls["btnConfiguraParametri"].Enabled = true;

                DataBase.CloseConnection();
            }
            else
            {
                //Se da default i controlli sono abilitati, li abilito
                if (Controls.IsDefaultEnabled("btnEsportaXML"))
                    Controls["btnEsportaXML"].Enabled = true;
                if (Controls.IsDefaultEnabled("btnImportaXML"))
                    Controls["btnImportaXML"].Enabled = true;

                //disabilito i controlli che non funzionano in emergenza
                if (Controls.Contains("btnProduzione"))
                    Controls["btnProduzione"].Enabled = false;
                if (Controls.Contains("btnTest"))
                    Controls["btnTest"].Enabled = false;
                if (Controls.Contains("btnDev"))
                    Controls["btnDev"].Enabled = false;
                if (Controls.Contains("btnAggiornaDati"))
                    Controls["btnAggiornaDati"].Enabled = false;
                if (Controls.Contains("btnAggiornaStruttura"))
                    Controls["btnAggiornaStruttura"].Enabled = false;
                if (Controls.Contains("btnConfiguraParametri"))
                    Controls["btnConfiguraParametri"].Enabled = false;
            }
        }
        /// <summary>
        /// Handler del click sul tasto di configurazione dei parametri. Apre il form che permette di modificare i valori dei parametri definiti per il foglio.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfiguraParametri_Click(object sender, RibbonControlEventArgs e)
        {
            FormModificaParametri form = new FormModificaParametri();
            if(!form.IsDisposed)
                form.Show();
        }
        /// <summary>
        /// Handler del SheetSelectionChange. Funzione che controlla se la cella selezionata è un Check. Si trova qui e non dentro la Classe Base.Handler perché deve interagire con l'errorPane 
        /// (non è possibile farlo dal namespace Base in quanto si creerebbe uno using circolare)
        /// </summary>
        /// <param name="Sh">Worksheet dove è stato eseguito il cambio di selezione</param>
        /// <param name="Target">Range dove è stato eseguito il cambio di selezione</param>
        private void CheckSelection(object Sh, Excel.Range Target)
        {
            try
            {
                if (!Workbook.FromErrorPane)
                {
                    DefinedNames definedNames = new DefinedNames(Target.Worksheet.Name, DefinedNames.InitType.Check);
                    Range rng = new Range(Target.Row, Target.Column);
                    if (definedNames.HasCheck() && definedNames.IsCheck(rng))
                    {
                        Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane = true;
                        _errorPane.SelectNode("'" + Target.Worksheet.Name + "'!" + rng.ToString());

                    }
                }
            }
            catch { }
        }
        /// <summary>
        /// Handler del ridimensionamento dell'ActionsPane del foglio, ridimensiona il componente ErrorPane in modo da adattarlo alle nuove dimensioni.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ActionsPane_SizeChanged(object sender, EventArgs e)
        {
            _errorPane.SetDimension(Globals.ThisWorkbook.ActionsPane.Width, Globals.ThisWorkbook.ActionsPane.Height);
        }
        /// <summary>
        /// Handler del click sui toggle buttons di cambio ambiente selezionato. Cambia la selezione, fa il refresh del file di configurazione e attiva l'aggiornamento della struttura del foglio.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelezionaAmbiente_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton ambienteScelto = (RibbonToggleButton)sender;

            if (DataBase.OpenConnection())
            {
                int count = 0;
                foreach (RibbonToggleButton button in FrontOffice.Groups.First(g => g.Label.ToLower() == "ambienti").Items)
                {
                    if (button.Checked)
                    {
                        button.Checked = false;
                        count++;
                    }
                }
                //se maggiore di 1 allora c'è un cambio ambiente altrimenti doppio click sullo stesso e non faccio nulla
                if (count > 1)
                {
                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogModifica, "Attivato ambiente " + ambienteScelto.Label);
                    DataBase.SwitchEnvironment(ambienteScelto.Label);

                    btnAggiornaStruttura_Click(null, null);
                }

                ambienteScelto.Checked = true;
                DataBase.CloseConnection();
            }
            else
            {
                ambienteScelto.Checked = false;

                System.Windows.Forms.MessageBox.Show("Non è possibile effettuare un cambio di ambiente quando il sistema è in emergenza...", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
            }
        }
        /// <summary>
        /// Handler del click del tasto di aggiornamento della struttura. Avvisa l'utente ed esegue l'aggiornamento della struttura. Esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAggiornaStruttura_Click(object sender, RibbonControlEventArgs e)
        {
            //avviso all'utente
            var response = System.Windows.Forms.DialogResult.Yes;
            
            if(sender != null && e != null)
                response = System.Windows.Forms.MessageBox.Show("Eseguire l'aggiornamento della struttura?", Simboli.NomeApplicazione + " - ATTENZIONE!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);

            if (response == System.Windows.Forms.DialogResult.Yes)
            {
                Workbook.ScreenUpdating = false;
                Sheet.Protected = false;

                Aggiorna aggiorna = new Aggiorna();
                if(aggiorna.Struttura(avoidRepositoryUpdate: false))
                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogModifica, "Aggiorna struttura");

                //29/01/2017 FIX errore caricamento quando numero mercati diverso tra un ambiente e l'altro (idem per stagioni)
                if (Controls.Contains("cmbMSD")) FillcmbMSD();
                if (Controls.Contains("cmbStagione")) FillcmbStagioni();
                //09/02/2017 MOD aggiunta combo mercati MI
                //if (Controls.Contains("MercatoMI")) FillcmbMI();
                
                RefreshChecks();

                Sheet.Protected = true;
                Workbook.ScreenUpdating = true;
                
                if (_allDisabled)
                    AbilitaTasti(true);
            }
        }
        /// <summary>
        /// Handler del click del tasto di cambio data. Verifica che la data selezionata sia diversa da quella attuale e fa partire il controllo per vedere se ci siano modifiche alla struttura attraverso
        /// DataBase.SP.CHECKMODIFICASTRUTTURA. Se ci sono aggiorno la struttra, altrimenti aggiorno semplicemente i dati. Esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCalendar_Click(object sender, RibbonControlEventArgs e)
        {
            //apro il form calendario
            FormCalendar cal = new FormCalendar();

            cal.Top = System.Windows.Forms.Cursor.Position.Y - 20;
            cal.Left = System.Windows.Forms.Cursor.Position.X - 20;

            DateTime calDate = cal.ShowDialog();
            cal.Dispose();
            Workbook.Application.Windows[1].Activate();
            //verifico che la data sia stata cambiata
            if (calDate != Workbook.DataAttiva)
            {
                //per validazione TL
                if (Workbook.IdApplicazione == 11 && calDate > DateTime.Today)
                    System.Windows.Forms.MessageBox.Show("La data selezionata nel calendario è successiva al giorno corrente.", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

                Workbook.ScreenUpdating = false;
                Sheet.Protected = false;
                SplashScreen.Show();

                ((RibbonButton)sender).Label = calDate.ToString("dddd dd MMM yyyy");

                Aggiorna aggiorna = new Aggiorna();
                if (DataBase.OpenConnection())
                {
                    Workbook.InsertLog(PSO.Core.DataBase.TipologiaLOG.LogModifica, "Cambio Data a " + ((RibbonButton)sender).Label);
                    DataBase.ExecuteSPApplicazioneInit(calDate);

                    bool aggiornaStruttura = CheckCambioStruttura(Workbook.DataAttiva, calDate);
                    
                    Workbook.DataAttiva = calDate;

                    if (aggiornaStruttura)
                        aggiorna.Struttura(avoidRepositoryUpdate: false);
                    else
                        aggiorna.Dati();

                    Workbook.RefreshLog();
                }
                else  //emergenza
                {
                    Workbook.DataAttiva = calDate;
                    aggiorna.Emergenza();
                }

                RefreshChecks();

                SplashScreen.Close();
                Sheet.Protected = true;
                Workbook.ScreenUpdating = true;
            }
        }
        /// <summary>
        /// Handler del click del tasto di selezione rampe. Apre il form per la selezione delle rampe ed esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRampe_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            //prendo il nome sheet e il range selezionato (per poter lavorare su più giorni nel caso ci fosse necessità)
            string sheet = Workbook.ActiveSheet.Name;
            Excel.Range rng = Workbook.Application.Selection;
            
            DefinedNames definedNames = new DefinedNames(sheet);
            FormSelezioneUP selUP = new FormSelezioneUP("PQNR_PROFILO");

            //controllo se nel range selezionato è definita un'entità
            if (sheet == "Iren Termo" && definedNames.IsDefined(rng.Row))
            {
                string nome = definedNames.GetNameByAddress(rng.Row, rng.Column);
                string siglaEntita = nome.Split(Simboli.UNION[0])[0];
                
                //controllo se l'entità ha la possibilità di selezionare le rampe
                DataView entitaInformazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
                entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaInformazione = 'PQNR_PROFILO' AND IdApplicazione = " + Workbook.IdApplicazione;

                if (entitaInformazioni.Count == 0)
                {
                    //avviso l'utente che l'entità selezionata non ha l'opzione
                    if (System.Windows.Forms.MessageBox.Show("L'operazione selezionata non è disponibile per l'UP selezionata, selezionarne un'altra dall'elenco?", Simboli.NomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes
                        && selUP.ShowDialog().ToString() != "")
                    {
                        Forms.FormRampe rampe = new FormRampe(Workbook.Application.Selection);
                        rampe.ShowDialog();
                        rampe.Dispose();
                    }
                }
                else
                {
                    Forms.FormRampe rampe = new FormRampe(Workbook.Application.Selection);
                    rampe.ShowDialog();
                    rampe.Dispose();
                }
            }
            //sono in un foglio diverso da Iren Termo o su una cella senza definizione di nomi
            else if (System.Windows.Forms.MessageBox.Show("Nessuna UP selezionata, selezionarne una dall'elenco?", Simboli.NomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes
                && selUP.ShowDialog().ToString() != "")
            {
                Forms.FormRampe rampe = new FormRampe(Workbook.Application.Selection);
                rampe.ShowDialog();
                rampe.Dispose();
            }
            selUP.Dispose();
            RefreshChecks();
            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler del click del tasto di aggiornamento dei dati. Aziona la funzione AggiornaDati ed esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAggiornaDati_Click(object sender, RibbonControlEventArgs e)
        { 
            var response = System.Windows.Forms.MessageBox.Show("Eseguire l'aggiornamento dei dati?", Simboli.NomeApplicazione + " - ATTENZIONE!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
            if (response == System.Windows.Forms.DialogResult.Yes)
            {
                Workbook.ScreenUpdating = false;
                Sheet.Protected = false;

                Aggiorna aggiorna = new Aggiorna();
                if (aggiorna.Dati())
                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogModifica, "Aggiorna Dati");

                RefreshChecks();

                Sheet.Protected = true;
                Workbook.ScreenUpdating = true;
            }
        }
        /// <summary>
        /// Handler del click del tasto delle azioni. Mostra il form delle azioni ed esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAzioni_Click(object sender, RibbonControlEventArgs e)
        {
            //aggiorno i limiti delle celle editabili
            if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1"))
            {
                Workbook.ScreenUpdating = false;
                Sheet.Protected = false;
                
                //04/02/2017 FIX: aggiornamento celle disabilitate all'avvio
                RefreshDisabledCells();

                Sheet.Protected = true;
                Workbook.ScreenUpdating = true;
            }

            FormAzioni frmAz = new FormAzioni(new Esporta(), new Riepilogo(), new Carica());
            frmAz.ShowDialog();

            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            RefreshChecks();

            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler del click del tasto di modifica. Attiva e disattiva la modifica foglio. Nel caso di disattivazione, aggiorna i check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnModifica_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.Application.EnableEvents = true;
            Workbook.ScreenUpdating = false;

            try
            {
                Sheet.Protected = false;
                Simboli.ModificaDati = ((RibbonToggleButton)sender).Checked;


                if (((RibbonToggleButton)sender).Checked)
                {
                    AbilitaTasti(false);
                    ((RibbonToggleButton)sender).Enabled = true;
                    ((RibbonToggleButton)sender).Image = PSO.Base.Properties.Resources.modificaSI;
                    ((RibbonToggleButton)sender).Label = "Modifica SI";
                    //Workbook.WB.SheetChange += Handler.StoreEdit;
                    Workbook.AddStdStoreEdit();
                    //Aggiungo handler per azioni custom nel caso servisse
                    Workbook.WB.SheetChange += _modificaCustom.Range;
                    
                    if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1"))
                    {
                        //04/02/2017 FIX: aggiornamento celle disabilitate all'avvio
                        RefreshDisabledCells();
                    }
                }
                else
                {
                    //salva modifiche sul db
                    Sheet.SalvaModifiche();
                    DataBase.SalvaModificheDB();
                    ((RibbonToggleButton)sender).Image = PSO.Base.Properties.Resources.modificaNO;
                    ((RibbonToggleButton)sender).Label = "Modifica NO";
                    //Workbook.WB.SheetChange -= Handler.StoreEdit;
                    Workbook.RemoveStdStoreEdit();
                    //Rimuovo handler per azioni custom nel caso servisse
                    Workbook.WB.SheetChange -= _modificaCustom.Range;

                    RefreshChecks();

                    //aggiorno i label dello stato nel caso sia necessario!
                    Workbook.AggiornaLabelStatoDB();

                    AbilitaTasti(true);
                    StatoDB_Changed(null, null);
                }
                Sheet.AbilitaModifica(((RibbonToggleButton)sender).Checked);

                Workbook.RefreshLog();
                Sheet.Protected = true;
            }
            catch(System.Runtime.InteropServices.COMException ex)
            {
                if (ex.Message.Contains("0x800A03EC"))
                {
                    System.Windows.Forms.MessageBox.Show("Prima di chiudere la modifica è necessario uscire dalla modalità di modifica della cella. Premere invio o selezionare un'altra cella.", Simboli.NomeApplicazione + " - ATTENZIONE!!!");
                }
                ((RibbonToggleButton)sender).Checked = !((RibbonToggleButton)sender).Checked;
            }

            Workbook.ScreenUpdating = true;
            
        }
        /// <summary>
        /// Handler del click del tasto di Ottimizzazione.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOttimizza_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            Excel.Range rng = Workbook.Application.Selection;

            DefinedNames definedNames = new DefinedNames(Workbook.ActiveSheet.Name, DefinedNames.InitType.Naming);

            //inizializzo ottimizzatore e il form di selezione entità per l'ottimo.
            Ottimizzatore opt = new Ottimizzatore();
            FormSelezioneUP selUP = new FormSelezioneUP("OTTIMO");

            if (definedNames.IsDefined(rng.Row))
            {
                string nome = definedNames.GetNameByAddress(rng.Row, rng.Column);
                string siglaEntita = nome.Split(Simboli.UNION[0])[0];

                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                
                if(categoriaEntita.Count > 0)
                    siglaEntita = categoriaEntita[0]["Gerarchia"] is DBNull ? siglaEntita : categoriaEntita[0]["Gerarchia"].ToString();

                DataView entitaInformazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
                entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaInformazione = 'OTTIMO' AND IdApplicazione = " + Workbook.IdApplicazione;

                if (entitaInformazioni.Count == 0)
                {
                    if(System.Windows.Forms.MessageBox.Show("L'UP selezionata non può essere ottimizzata, selezionarne un'altra dall'elenco?", Simboli.NomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes)
                    {
                        siglaEntita = selUP.ShowDialog().ToString();
                        if (siglaEntita != null)
                        {
                            SplashScreen.Show();
                            SplashScreen.UpdateStatus("Ottimizzo " + siglaEntita);
                            opt.EseguiOttimizzazione(siglaEntita);
                            SplashScreen.UpdateStatus("Salvo modifiche su DB");
                            Sheet.SalvaModifiche();
                            DataBase.SalvaModificheDB();
                            SplashScreen.Close();
                        }
                    }
                }
                else
                {
                    SplashScreen.Show();
                    SplashScreen.UpdateStatus("Ottimizzo " + siglaEntita);
                    opt.EseguiOttimizzazione(siglaEntita);
                    SplashScreen.UpdateStatus("Salvo modifiche su DB");
                    Sheet.SalvaModifiche();
                    DataBase.SalvaModificheDB();
                    SplashScreen.Close();
                }
            }
            else if (System.Windows.Forms.MessageBox.Show("Nessuna UP selezionata, selezionarne una dall'elenco?", Simboli.NomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes)
            {
                object siglaEntita = selUP.ShowDialog();
                if (siglaEntita != null)
                {
                    SplashScreen.Show();
                    SplashScreen.UpdateStatus("Ottimizzo " + siglaEntita);
                    opt.EseguiOttimizzazione(siglaEntita);
                    SplashScreen.UpdateStatus("Salvo modifiche su DB");
                    Sheet.SalvaModifiche();
                    DataBase.SalvaModificheDB();
                    SplashScreen.Close();
                }
            }
            selUP.Dispose();
            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler del click del tasto di modifica parametri. Mostra il form di modifica dei parametri utente.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfigura_Click(object sender, RibbonControlEventArgs e)
        {
            FormConfiguraPercorsi conf = new FormConfiguraPercorsi();
            conf.ShowDialog();
            conf.Dispose();
        }
        /// <summary>
        /// Handler del click dei tasti delle varie applicazioni. Abilita il foglio selezionato.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnProgrammi_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton btn = (RibbonToggleButton)sender;

            if (!btn.Checked)
            {
                btn.Checked = true;
            }
            else
            {
                btn.Checked = false;
                //TODO Controllare e cambiare il path
                //Handler.SwitchWorksheet(btn.Name.Substring(3));

                Workbook.AvviaApplicazione(Workbook.Application, (int)btn.Tag);
            }
        }
        /// <summary>
        /// Handler del click del tasto per forzare l'emergenza. Disabilita le connessioni al DB.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnForzaEmergenza_Click(object sender, RibbonControlEventArgs e)
        {
            Simboli.EmergenzaForzata = ((RibbonToggleButton)sender).Checked;
            StatoDB_Changed(null, null);
        }
        /// <summary>
        /// Handler del click del tasto di chiusura. Chiude l'applicativo.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChiudi_Click(object sender, RibbonControlEventArgs e)
        {
            TextInfo ti = new CultureInfo("it-IT", false).TextInfo;
            string pathStr = Workbook.Repository.Applicazione["PathBackup"].ToString();
            try
            {
                if (Workbook.Ambiente == Simboli.PROD)
                {
                    if (!Directory.Exists(pathStr))
                        Directory.CreateDirectory(pathStr);

                    string filename = ti.ToTitleCase(Simboli.NomeApplicazione).Replace(" ", "") + "_Backup_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_v" + Workbook.WorkbookVersion.ToString() + ".xlsm";

                    Globals.ThisWorkbook.SaveCopyAs(Path.Combine(pathStr, filename));
                }

                //02/02/2017 FIX: provo a terminare excel alla chiusura
                //cerca la finestra ed invia il segnale di chiusura
                //int WM_CLOSE = 16;
                //IntPtr HWnd = FindWindow(null, "Microsoft Excel - " + Globals.ThisWorkbook.Name);
                //PostMessage(HWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);

                Globals.ThisWorkbook.Close(true, Type.Missing, Type.Missing);
            }
            catch(DirectoryNotFoundException)
            {
                if(System.Windows.Forms.MessageBox.Show("Il percorso di backup non è raggiungibile. Chiudere comunque il file senza eseguire il backup?", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
                    Globals.ThisWorkbook.Close(true, Type.Missing, Type.Missing);
            }
        }
        /// <summary>
        /// Handler del click del tasto per visualizzare l'actionsPane del documento.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMostraErrorPane_Click(object sender, RibbonControlEventArgs e)
        {            
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            if (!Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane)
                Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane = true;
            
            _errorPane.SetDimension(Globals.ThisWorkbook.ActionsPane.Width, Globals.ThisWorkbook.ActionsPane.Height);
            
            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler della selezione di un nuovo mercato in cmbMSD su ribbon. Aggiorna la struttura dei fogli.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbMSD_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            Workbook.IdApplicazione = ((RibbonDropDown)sender).SelectedItem.Tag;
            
            Aggiorna aggiorna = new Aggiorna();
            aggiorna.Struttura(avoidRepositoryUpdate: true);

            RefreshChecks();

            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler per il cambio di stagione da cmnStagione su ribbon. Imposta il valore della riga nascosta.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbStagione_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;
            
            Workbook.IdStagione = ((RibbonDropDown)sender).SelectedItem.Tag;
            
            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }

        private void btnEsportaXML_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            SplashScreen.Show();
            SplashScreen.UpdateStatus("Esporto tutte le informazioni del foglio");

            EsportaXML exp = new EsportaXML();
            exp.RunExport();
            SplashScreen.Close();
            Workbook.ScreenUpdating = true;
        }

        private void btnImportaXML_Click(object sender, RibbonControlEventArgs e)
        {
            FormImportXML frmXML = new FormImportXML();
            frmXML.ShowDialog();
        }

        private void btnFormIncremento_Click(object sender, RibbonControlEventArgs e)
        {
            //TODO riportare ad originale, ora intercetto l'evento per attivare il form di bilanciamento
            
            Excel.Worksheet ws = Workbook.Application.ActiveSheet;
            if(!Workbook.CategorySheets.Contains(ws))
            {
                System.Windows.Forms.MessageBox.Show("Nel foglio selezionato non è possibile eseguire questa operazione.", Simboli.NomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);

                return;
            }

            Win32Window wnd = new Win32Window(new IntPtr(Workbook.Application.Hwnd));

            if (Workbook.IdApplicazione == 18)
            {
                if ((_frmIncMI == null || _frmIncMI.IsDisposed))
                {
                    _frmIncMI = new FormIncrementoMI(ws, Workbook.Application.Selection);
                    _frmIncMI.FormClosed += FrmInc_FormClosed;

                    _frmIncMI.Show(wnd);
                }
            }
            else
            {
                if ((_frmInc == null || _frmInc.IsDisposed))
                {
                _frmInc = new FormIncremento(ws, Workbook.Application.Selection);
                _frmInc.FormClosed += FrmInc_FormClosed;

                    _frmInc.Show(wnd);
                }
            }

            wnd = null;
            
            AbilitaTasti(false, "btnIncremento");
        }

        private void btnCancelOffers_Click(object sender, RibbonControlEventArgs e)
        {
            DataView categoria = Workbook.Repository[DataBase.TAB.CATEGORIA].DefaultView;
            categoria.RowFilter = "DesCategoria = '" + Workbook.ActiveSheet.Name + "'";

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoria.RowFilter = "DesCategoria = '" + Workbook.ActiveSheet.Name + "'";

            Excel.Range Target = Workbook.Application.Selection;

            Excel.Worksheet ws = Workbook.Application.ActiveSheet;

            DefinedNames _definedNames = new DefinedNames(ws.Name, DefinedNames.InitType.All);

            if (Target.Rows.Count > 1)
            {
                foreach (Excel.Range row in Target.Rows)
                {
                    if (row.EntireRow.Hidden)
                    {
                        System.Windows.Forms.MessageBox.Show("ERRORE: Nel range selezionato ci sono righe nascoste.");
                        return;
                    }
                }

                System.Windows.Forms.MessageBox.Show("ATTENZIONE: Nel range selezionato ci sono più righe.");
                return;
            }
            foreach (Excel.Range row in Target.Rows)
            {
                if (!_definedNames.IsEditable(row.Row))
                {
                    System.Windows.Forms.MessageBox.Show("ERRORE: Il range selezionato contiene delle righe non modificabili.");
                    return;
                }
            }

            int firstCol = _definedNames.GetFirstCol();
            if (Target.Column < firstCol + Simboli.GetMarketOffsetMI(Workbook.Mercato, Workbook.DataAttiva))
            {
                System.Windows.Forms.MessageBox.Show("ERRORE: Il range selezionato contiene celle appartenenti a mercati chiusi.");
                return;
            }

            string name_selected_row = _definedNames.GetNameByRow(Target.Row).FirstOrDefault();
            string name_associated_row = "";

            string pattern = @"^Offerta(.*)P$";
            string[] splitted = name_selected_row.Split(Simboli.UNION[0]);

            //Ottengo tutte le altre siglaEntità escusa quella su cui si agisce
            // mi servirà per ciclare su tutti gli altri gradini possibili per un eventuale bilanciamento
            categoriaEntita.RowFilter = "SiglaCategoria = '" + categoria[0]["SiglaCategoria"] + "' AND Gerarchia is null AND SiglaEntita <> '" + splitted[0] + "'";

            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
            bool is_price = rgx.IsMatch(splitted[splitted.Length - 1]);
            if (is_price)
            {
                name_associated_row = name_selected_row.Substring(0, name_selected_row.Length - 1) + "E";
            }
      
            pattern = @"^Offerta(.*)E$";
            rgx = new Regex(pattern, RegexOptions.IgnoreCase);
            bool is_quantity = rgx.IsMatch(splitted[splitted.Length - 1]);
            if (is_quantity)
            {
                name_associated_row = name_selected_row.Substring(0, name_selected_row.Length - 1) + "P";
            }

            if (!is_price && !is_quantity)
            {
                System.Windows.Forms.MessageBox.Show("ERRORE: Per cancellare un'offerta selezionare singole righe prezzo/energia.");
                return;
            }

            string name_codBil_row = name_selected_row.Substring(0, name_selected_row.Length - 1) + "CB";
            string name_type_row = name_selected_row.Substring(0, name_selected_row.Length - 1) + "TIPO";

            //int selected_row = _definedNames.GetRowByName(name_selected_row);
            int associated_row = _definedNames.GetRowByName(name_associated_row);
            int type_row = _definedNames.GetRowByName(name_type_row);
            int codBil_row = _definedNames.GetRowByName(name_codBil_row);

            Sheet.Protected = false;

            if (Workbook.Repository[DataBase.TAB.MODIFICA] == null)
                Workbook.Repository.Add(Workbook.Repository.CreaTabellaModifica(DataBase.TAB.MODIFICA));

            foreach (Excel.Range r in Target)
            {
                r.Value2 = null;
                ws.Cells[associated_row, r.Column].Value2 = null;
                ws.Cells[type_row, r.Column].Value2 = "VEN";

                Handler.StoreEdit(r, 0);
                Handler.StoreEdit(ws.Cells[associated_row, r.Column], 0);
                Handler.StoreEdit(ws.Cells[type_row, r.Column], 0);

                if ( !(ws.Cells[codBil_row, r.Column].Value2 == null || ws.Cells[codBil_row, r.Column].Value2.ToString() == "") )
                {
                    foreach (DataRowView drv in categoriaEntita)
                    {
                        //
                        for (int i = 1; i < 5; i++)
                        {
                            string name_row_type_tmp = drv["SiglaEntita"] + Simboli.UNION +
                                splitted[1].Substring(0, splitted[1].Length - 2);

                            int row_cod_tmp = _definedNames.GetRowByName(name_row_type_tmp + i + "CB");
                            int row_q_tmp = _definedNames.GetRowByName(name_row_type_tmp + i + "E");
                            int row_p_tmp = _definedNames.GetRowByName(name_row_type_tmp + i + "P");
                            int row_type_tmp = _definedNames.GetRowByName(name_row_type_tmp + i + "TIPO");

                            // Cancello solo i CodBilanciamento uguali all'offerta selezionata
                            if (!(ws.Cells[row_cod_tmp, r.Column].Value2 == null || ws.Cells[row_cod_tmp, r.Column].Value2.ToString() == "") && ws.Cells[row_cod_tmp, r.Column].Value2 == ws.Cells[codBil_row, r.Column].Value2)
                            {
                                ws.Cells[row_cod_tmp, r.Column].Value2 = null;
                                ws.Cells[row_q_tmp, r.Column].Value2 = null;
                                ws.Cells[row_p_tmp, r.Column].Value2 = null;
                                ws.Cells[row_type_tmp, r.Column].Value2 = "VEN";

                                Handler.StoreEdit(ws.Cells[row_cod_tmp, r.Column], 0);
                                Handler.StoreEdit(ws.Cells[row_q_tmp, r.Column], 0);
                                Handler.StoreEdit(ws.Cells[row_p_tmp, r.Column], 0);
                                Handler.StoreEdit(ws.Cells[row_type_tmp, r.Column], 0);
                            }
                        }
                    }

                    ws.Cells[codBil_row, r.Column].Value2 = null;
                    Handler.StoreEdit(ws.Cells[codBil_row, r.Column], 0);
                }
            }

            Handler.StoreEdit(Target);

            //Cancello i commenti dopo la Store perchè altrimenti verrebbero rivisualizzati in redraw
            foreach (Excel.Range r in Target)
            {
                r.ClearComments();
                ws.Cells[associated_row, r.Column].ClearComments();
                ws.Cells[type_row, r.Column].ClearComments();
            }

            DataBase.SalvaModificheDB(DataBase.TAB.MODIFICA);
            Workbook.Repository.Remove(DataBase.TAB.MODIFICA);

            Sheet.Protected = true;
        }

        private void FrmInc_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            AbilitaTasti(true);
            StatoDB_Changed(null, null);

            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            RefreshChecks();
            
            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;

            Workbook.Application.Visible = true;
            
        }

        private void FrmBlc_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            AbilitaTasti(true);
            StatoDB_Changed(null, null);

            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            RefreshChecks();

            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;

            Workbook.Application.Visible = true;

        }

        private void cmbMI_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            Workbook.Mercato = ((RibbonDropDown)Controls["MercatoMI"]).SelectedItem.Label;

            Aggiorna aggiorna = new Aggiorna();
            aggiorna.Dati();

            RefreshChecks();

            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }

        //19/02/2017 MOD: inserito form bilanciamento
        private void btnOfferteBilanciate_Click(object sender, RibbonControlEventArgs e) 
        {
            Excel.Worksheet ws = Workbook.Application.ActiveSheet;
            if (!Workbook.CategorySheets.Contains(ws))
            {
                System.Windows.Forms.MessageBox.Show("Nel foglio selezionato non è possibile eseguire questa operazione.", Simboli.NomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);

                return;
            }

            Win32Window wnd = new Win32Window(new IntPtr(Workbook.Application.Hwnd));

            if (Workbook.IdApplicazione == 18)
            {
                if ((_frmBalance == null || _frmBalance.IsDisposed))
                {
                    _frmBalance = new FormBilanciamento(ws, Workbook.Application.Selection);
                    _frmBalance.FormClosed += FrmBlc_FormClosed;

                    _frmBalance.Show(wnd);
                }
            }


            wnd = null;
        }
        
        #endregion

        #region Metodi

        /// <summary>
        /// Funzione per aggiornate i check in seguito ad operazioni di modifica del foglio.
        /// </summary>
        private void RefreshChecks()
        {
            SplashScreen.Show();
            Workbook.ScreenUpdating = false;

            bool autoCalc = Workbook.Application.Calculation == Excel.XlCalculation.xlCalculationAutomatic;
            if(autoCalc)
                Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            try
            {
                System.Diagnostics.Stopwatch wc = System.Diagnostics.Stopwatch.StartNew();
                _errorPane.Clear();
                _errorPane.RefreshCheck(_checkFunctions);
                wc.Stop();
                //quando non ci sono funzioni di check rischia di andare in errore perché non riesce ad inizializzare la splashscreen prima di chiuderla
                if(wc.ElapsedMilliseconds < 3)
                    System.Threading.Thread.Sleep(3);
            }
            catch { }
            SplashScreen.Close();
            
            if(autoCalc)
                Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        }

        //04/02/2017 FIX: Aggiornamento celle disabilitate all'avvio
        private void RefreshDisabledCells()
        {
            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.MakeCellsDisabled();
                s.UpdateDayColor();
            }
        }
        /// <summary>
        /// Metodo di inizializzazione della Tab Front Office. Visualizza e abilita i tasti in base alle specifiche da DB. Da notare che se ci sono aggiornamenti, bisogna caricare la struttura e riavviare l'applicativo.
        /// </summary>
        private void Initialize()
        {
//            _controls = new ControlCollection(this);
//            DataView controlli = new DataView();
            
//            if (DataBase.OpenConnection())
//            {
//                Workbook.Repository.CaricaApplicazioneRibbon();
//                controlli = Workbook.Repository[DataBase.TAB.APPLICAZIONE_RIBBON].DefaultView;
//                DataBase.CloseConnection();
//            }
//            else
//            {
//                try
//                {
//                    controlli = Workbook.Repository[DataBase.TAB.APPLICAZIONE_RIBBON].DefaultView;
//                }
//                catch
//                {
//                    controlli = new DataView();
//                }
//            }

//            if (controlli.Count > 0)
//            {
//                foreach (DataRowView controllo in controlli)
//                {
//                    Controls[controllo["NomeControllo"].ToString()].Visible = controllo["Visibile"].Equals("1");
//                    Controls[controllo["NomeControllo"].ToString()].Enabled = controllo["Abilitato"].Equals("1");
//                    if (controllo["Abilitato"].Equals("1"))
//                        _enabledControls.Add(controllo["NomeControllo"].ToString());

//                    if (Controls[controllo["NomeControllo"].ToString()].GetType().ToString().Contains("ToggleButton"))
//                    {
//                        ((RibbonToggleButton)Controls[controllo["NomeControllo"].ToString()]).Checked = controllo["Stato"].Equals("1");
//                    }
//                }

//                List<RibbonGroup> groups = FrontOffice.Groups.ToList();
//                foreach (RibbonGroup group in groups)
//                    group.Visible = group.Items.Any(c => c.Visible);
//            }
//            else
//            {
//                foreach (RibbonControl control in Controls)
//                {
//#if !DEBUG
//                    control.Visible = true;
//                    control.Enabled = false;
//#else
//                    control.Visible = true;
//                    control.Enabled = true;
//#endif

//                    if (control.GetType().ToString().Contains("ToggleButton"))
//                        ((RibbonToggleButton)control).Checked = false;
//                }
//            }

            //ComboBox mercati
            
        }
        /// <summary>
        /// Metodo che seleziona il tasto corretto tra quelli degli applicativi presenti nella Tab Front Office. La selezione avviene in base all'ID applicazione scritto sul file di configurazione.
        /// </summary>
        private void CheckTastoApplicativo()
        {
            switch (ConfigurationManager.AppSettings["AppID"])
            {
                case "1":
                    ((RibbonToggleButton)Controls["btnOfferteMGP"]).Checked = true;
                    break;
                case "2":
                case "3":
                case "4":
                case "13":
                    ((RibbonToggleButton)Controls["btnInvioProgrammi"]).Checked = true;
                    break;
                case "5":
                    ((RibbonToggleButton)Controls["btnProgrammazioneImpianti"]).Checked = true;
                    break;
                case "6":
                    ((RibbonToggleButton)Controls["btnUnitCommitment"]).Checked = true;
                    break;
                case "7":
                    ((RibbonToggleButton)Controls["btnPrezziMSD"]).Checked = true;
                    break;
                case "8":
                    ((RibbonToggleButton)Controls["btnSistemaComandi"]).Checked = true;
                    break;
                case "9":
                    ((RibbonToggleButton)Controls["btnOfferteMSD"]).Checked = true;
                    break;
                case "10":
                    ((RibbonToggleButton)Controls["btnOfferteMB"]).Checked = true;
                    break;
                case "11":
                    ((RibbonToggleButton)Controls["btnValidazioneTL"]).Checked = true;
                    break;
                case "12":
                    ((RibbonToggleButton)Controls["btnPrevisioneCT"]).Checked = true;
                    break;
            }



        }
        /// <summary>
        /// Abilito tutti i tasti nel caso in cui, ad esempio in seguito a un rilascio, questi vengano disabilitati da DisabilitaTasti.
        /// </summary>
        private void AbilitaTasti(bool enable, params string[] exceptions)
        {
            foreach (RibbonControl control in Controls.GetDefaultEnabled())
                if(!Array.Exists(exceptions, s => s == control.Name))
                    control.Enabled = enable;

            _allDisabled = enable;
        }
        /// <summary>
        /// Imposta il mercato attivo in base all'orario. Se necessario cambia anche la data attiva e imposta il foglio come da aggiornare.
        /// </summary>
        /// <param name="appID">L'ID applicazione che identifica anche in quale mercato il foglio è impostato.</param>
        /// <param name="dataAttiva">La data attiva da modificare all'occorrenza.</param>
        /// <returns>Restituisce true se il foglio è da aggiornare, false altrimenti.</returns>
        private void SetMercato(out int newAppId)
        {
            //configuro la data attiva
            int ora = DateTime.Now.Hour;
            ((RibbonDropDown)Controls["cmbMSD"]).SelectedItem = ((RibbonDropDown)Controls["cmbMSD"]).Items.Where(i => i.Label == Simboli.OreMSD[ora]).First();
            ((RibbonDropDown)Controls["cmbMSD"]).SelectionChanged += cmbMSD_SelectionChanged;

            newAppId = ((RibbonDropDown)Controls["cmbMSD"]).SelectedItem.Tag;
        }

        //09/02/2017 MOD: Gestione mercati MI
        private void SetMercatoMI()
        {
            //configuro la data attiva
            int ora = DateTime.Now.Hour;
            ((RibbonDropDown)Controls["MercatoMI"]).SelectedItem = ((RibbonDropDown)Controls["MercatoMI"]).Items.Where(i => i.Label == Simboli.GetActiveMarket(ora)).First();
            ((RibbonDropDown)Controls["MercatoMI"]).SelectionChanged += cmbMI_SelectionChanged;

            Workbook.Mercato = ((RibbonDropDown)Controls["MercatoMI"]).SelectedItem.Label;
        }

        /// <summary>
        /// Aggiorna la data per le applicazione Validazione TL e Previsione CT.
        /// </summary>
        /// <param name="appID">L'ID applicazione</param>
        /// <param name="dataAttiva">La data attiva da cambiare se necessario</param>
        /// <returns>Restituisce true se il foglio è da aggiornare, false altrimenti.</returns>
        private void AggiornaData(out DateTime newDate, out bool refused)
        {
            bool done = true;
            refused = false;
            newDate = Workbook.DataAttiva;
            int ora = DateTime.Now.Hour;
            switch (Workbook.IdApplicazione)
            {
                case 1:
                    if (ora < 14)
                        newDate = DateTime.Today.AddDays(1);
                    else
                        newDate = DateTime.Today.AddDays(2);
                    break;
                //FIX aggiunti MSD5/6 e mandato avanti orario
                case 2:
                case 3:
                case 4:
                case 13:
                case 16:
                case 17:
                    if (ora > 19)
                        newDate = DateTime.Today.AddDays(1);
                    else
                        newDate = DateTime.Today;
                    break;
                case 5:
                    newDate = DateTime.Today;
                    break;
                case 6:
                    newDate = DateTime.Today.AddDays(-5);
                    break;
                case 7:
                    newDate = DateTime.Today.AddDays(-3);
                    break;
                case 8:
                case 9:
                    if (ora < 12)
                        newDate = DateTime.Today;
                    else
                        newDate = DateTime.Today.AddDays(1);
                    break;
                case 10:
                    if (ora < 22)
                        newDate = DateTime.Today;
                    else
                        newDate = DateTime.Today.AddDays(1);
                    break;
                case 11:
                    newDate = DateTime.Today.AddDays(-1);
                    break;
                case 12:
                    if (ora <= 15)
                        newDate = DateTime.Today.AddDays(1);
                    else
                        newDate = DateTime.Today.AddDays(2);
                    break;
                case 14:
                    DataTable dt = DataBase.Select(DataBase.SP.GET_LAST_DATA_VALIDATA_GAS);
                    if (dt != null && dt.Rows.Count > 0)
                        newDate = DateTime.ParseExact(dt.Rows[0][0].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                    break;
                case 15:
                    newDate = DateTime.Today;
                    break;
                //06/02/2017 MOD: Aggiunto cambio giorno mercati MI
                case 18:
                    if (ora >= 13)
                        newDate = DateTime.Today.AddDays(1);
                    else
                        newDate = DateTime.Today;
                    break;
                default:
                    done = false;
                    break;
            }

            if (done)
            {
                //25/01/2017 FIX: Cambio data da console non funzionava
                if (!Workbook.DaConsole || (Workbook.DaConsole && !Workbook.AccettaCambioData))
                {
                    if (newDate != Workbook.DataAttiva)
                    {
                        SplashScreen.Close();
                        refused = System.Windows.Forms.MessageBox.Show("La data sta per essere cambiata in " + newDate.ToString("ddd dd MMM yyyy") + ".\nIl cambiamento della data comporta un aggiornamento di tutte le informazioni. Vuoi continuare?", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Information) == System.Windows.Forms.DialogResult.No;
                        SplashScreen.Show();
                        if (refused)
                            newDate = Workbook.DataAttiva;
                    }
                }
            }
        }

        private bool CheckCambioStruttura(DateTime vecchia, DateTime nuova)
        {
            //controllo se nell'intervallo ci sono giorni a 23 o 25 ore
            //vecchia + intervallo
            //nuova + intervallo
            int intervalloGiorniMax = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA]
                .AsEnumerable()
                .Where(r => r["IdApplicazione"].Equals(Workbook.IdApplicazione))
                .Where(r => r["SiglaProprieta"].ToString().EndsWith("GIORNI_STRUTTURA"))
                .Select(r => int.Parse(r["Valore"].ToString()))
                .DefaultIfEmpty()
                .Max();
            
            intervalloGiorniMax = Math.Max(intervalloGiorniMax, Struct.intervalloGiorni);

            int oldPeriodHours = Date.GetOreIntervallo(vecchia, vecchia.AddDays(intervalloGiorniMax));
            int newPeriodHours = Date.GetOreIntervallo(nuova, nuova.AddDays(intervalloGiorniMax));

            if (oldPeriodHours != 24 * (intervalloGiorniMax + 1) || newPeriodHours != 24 * (intervalloGiorniMax + 1))
                return true;

            DataTable stato = DataBase.Select(DataBase.SP.CHECKMODIFICASTRUTTURA, "@DataOld=" + vecchia.ToString("yyyyMMdd") + ";@DataNew=" + nuova.ToString("yyyyMMdd"));

            return stato != null && stato.Rows.Count > 0 && stato.Rows[0]["Stato"].Equals(1);
        }

        public void InitRibbon()
        {
            if (!Workbook.AbortedLoading)
            {
                Sheet.Protected = false;
                Workbook.ScreenUpdating = false;
                SplashScreen.Show();

                //salvo gli eventuali valori aggiornati negli oggetti cached
                if (_updated)
                {
                    Globals.ThisWorkbook.ribbonDataSet.Tables.Clear();
                    Globals.ThisWorkbook.ribbonDataSet.Tables.Add(_dtControllo);
                    Globals.ThisWorkbook.ribbonDataSet.Tables.Add(_dtControlloApplicazione);
                    Globals.ThisWorkbook.ribbonDataSet.Tables.Add(_dtFunzioni);
                }

                //Seleziono l'ambiente in funzione dei tasti attivi nel menu
                switch (Workbook.Ambiente)
                {
                    case Simboli.PROD:
                        if (!Controls["btnProd"].Enabled && !Controls["btnTest"].Enabled)
                            Workbook.Ambiente = Simboli.DEV;
                        else if (!Controls["btnProd"].Enabled)
                            Workbook.Ambiente = Simboli.TEST;
                        break;
                    case Simboli.TEST:
                        if (!Controls["btnTest"].Enabled)
                            Workbook.Ambiente = Simboli.DEV;
                        break;
                }

                ((RibbonToggleButton)Controls["btn" + Workbook.Ambiente]).Checked = true;
                bool switchEnv = DataBase.SwitchEnvironment(Workbook.Ambiente);

                Workbook.StatoDBChanged(null, null);

                if (Controls.Contains("cmbMSD")) FillcmbMSD();
                if (Controls.Contains("cmbStagione")) FillcmbStagioni();
                //09/02/2017 MOD aggiunta combo mercati MI
                if (Controls.Contains("MercatoMI")) FillcmbMI();

                //per Invio Programmi
                DateTime newDate = Workbook.DataAttiva;
                int newIdApplicazione = Workbook.IdApplicazione;

                if (Controls.Contains("cmbMSD")) SetMercato(out newIdApplicazione);
                
                if (Controls.Contains("MercatoMI")) SetMercatoMI();
                

                Riepilogo r = new Riepilogo(Workbook.Main);

                Aggiorna aggiorna = new Aggiorna();
                bool aggiornaStruttura = false;
                

                if (DataBase.OpenConnection())
                {
                    //25/01/2017 FIX: Cambio data da console non funzionava
                    bool rifiutatoCambioData = Workbook.RifiutaCambioData;
                    if(!Workbook.DaConsole || (Workbook.DaConsole && !Workbook.RifiutaCambioData))
                        AggiornaData(out newDate, out rifiutatoCambioData);

                    bool aggiornaDati = Workbook.DataAttiva != newDate || Workbook.AggiornaDati;

                    aggiornaStruttura = CheckCambioStruttura(Workbook.DataAttiva, newDate) || Workbook.IdApplicazione != newIdApplicazione || Workbook.DaAggiornare;

                    if (Workbook.DataAttiva != newDate)
                        DataBase.ExecuteSPApplicazioneInit(newDate);

                    Workbook.DataAttiva = newDate;
                    Workbook.IdApplicazione = newIdApplicazione;

                    if (aggiornaStruttura)
                        aggiorna.Struttura(avoidRepositoryUpdate: !switchEnv);
                    else if(!rifiutatoCambioData)
                    {
                        if (aggiornaDati)
                            aggiorna.Dati();
                        else 
                        {
                            SplashScreen.Close();
                            bool scelta = Workbook.AggiornaDati;
                            if (!Workbook.DaConsole)
                            {
                                scelta = System.Windows.Forms.MessageBox.Show("Aggiornare i dati?", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes;
                            }
                            SplashScreen.Show();
                            if (scelta)
                                aggiorna.Dati();
                        }
                    }
                }
                else
                {
                    SplashScreen.Close();
                    System.Windows.Forms.MessageBox.Show("Il file si è aperto in condizioni di emergenza. I dati non sono aggiornati. La data può essere modificata manualmente.", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    SplashScreen.Show();
                    r.RiepilogoInEmergenza();
                }

                if (!aggiornaStruttura)
                    aggiorna.ColoriData();

                r.InitLabels();
                ((RibbonButton)Controls["btnCalendario"]).Label = Workbook.DataAttiva.ToString("dddd dd MMM yyyy");

                //aggiungo errorPane
                if (Controls.Contains("btnMostraErrorPane"))
                {
                    Globals.ThisWorkbook.ActionsPane.Controls.Add(_errorPane);
                    Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane = false;
                    Globals.ThisWorkbook.ActionsPane.AutoScroll = false;
                    Globals.ThisWorkbook.ActionsPane.SizeChanged += ActionsPane_SizeChanged;
                }

                //04/02/2017 FIX: aggiornamento celle disabilitate all'avvio
                if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1"))
                {
                    RefreshDisabledCells();
                }

                RefreshChecks();

                try
                {
                    Sheet.AbilitaModifica(false);
                }
                catch { }

                //aggiungo un altro handler per cell click
                Globals.ThisWorkbook.SheetSelectionChange += CheckSelection;
                Globals.ThisWorkbook.SheetSelectionChange += Handler.SelectionClick;

                //aggiungo un handler per modificare lo stato dei tasti di export a seconda dello stato del DB
                StatoDB_Changed(null, null);

                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogAccesso, "Log on - " + Environment.UserName + " - " + Environment.MachineName);

                Sheet.Protected = true;
                SplashScreen.Close();


                //TODO inserire qui i metodi per fare il carica/genera/esporta automatizzato quando si esegue da console
                //sarebbe auspicabile avere delle azioni per cui non serva passare la lista di entita o le date su cui operare
                //chiedere a DOMENICO
                if (Workbook.DaConsole && Workbook.HaAzioni && Workbook.HaEntita)
                {
                    FormAzioni frmAz = new FormAzioni(new Esporta(), new Riepilogo(), new Carica());
                    frmAz.ShowDialog(true, Workbook.ListaAzioni, Workbook.ListaEntita);

                    //19/01/2017 FIX DomainUloadedException
                    //spostata chiusura in ToolsExcelRibbon_Load()
                }
            }
        }

        #endregion
    }

    #region Controls Collection

    /// <summary>
    /// Classi che permettono di indicizzare per nome tutti i controlli contenuti nei gruppi della Tab Front Office
    /// </summary>
    public class ControlCollection : IEnumerable
    {
        #region Variabili

        private ToolsExcelRibbon _ribbon;
        private Dictionary<string, RibbonControl> _controls = new Dictionary<string, RibbonControl>();
        private Dictionary<string, bool> _defaultEnabled = new Dictionary<string, bool>();

        #endregion

        #region Proprietà

        public int Count
        {
            get { return _controls.Count; }
        }

        public RibbonControl this[int i]
        {
            get { return _controls.ElementAt(i).Value; }
        }

        public RibbonControl this[string name]
        {
            get { return _controls[name]; }
        }

        #endregion

        #region Costruttori

        internal ControlCollection(ToolsExcelRibbon ribbon)
        {
            _ribbon = ribbon;
        }

        #endregion

        #region Metodi

        public void Add(RibbonControl control)
        {
            _controls.Add(control.Name, control);
            _defaultEnabled.Add(control.Name, control.Enabled);
        }

        public bool Contains(string name)
        {
            return _controls.ContainsKey(name);
        }

        public bool IsDefaultEnabled(string name)
        {
            if (_defaultEnabled.ContainsKey(name))
                return _defaultEnabled[name];

            return false;
        }

        public IEnumerable<RibbonControl> GetDefaultEnabled()
        {
            return
                from kv in _controls
                where _defaultEnabled[kv.Key]
                select kv.Value;
        }

        public IEnumerator GetEnumerator()
        {
            return new ControlEnumerator(_ribbon);
        }

        public IEnumerable<KeyValuePair<string, RibbonControl>> AsEnumerable()
        {
            return _controls.AsEnumerable();
        }

        #endregion
    }
    public class ControlEnumerator : IEnumerator
    {
        #region Variabili

        private ToolsExcelRibbon _ribbon;
        private int _pos = -1;
        private int _max = -1;

        #endregion

        #region Costruttori

        public ControlEnumerator(ToolsExcelRibbon ribbon)
        {
            _ribbon = ribbon;
            _max = ribbon.Controls.Count;
        }

        #endregion

        #region Metodi

        public object Current
        {
            get { return _ribbon.Controls[_pos]; }
        }
        public bool MoveNext()
        {
            _pos++;
            return _pos < _max;
        }
        public void Reset()
        {
            _pos = -1;
        }

        #endregion
    }

    #endregion

    #region Groups Collection

    /// <summary>
    /// Classi che permettono di indicizzare per nome tutti i gruppi contenuti nei gruppi della Tab Front Office
    /// </summary>
    public class GroupsCollection : IEnumerable
    {
        #region Variabili

        private ToolsExcelRibbon _ribbon;
        private Dictionary<string, RibbonGroup> _groups = new Dictionary<string, RibbonGroup>();

        #endregion

        #region Proprietà

        public int Count
        {
            get { return _groups.Count; }
        }

        public RibbonGroup this[int i]
        {
            get { return _groups.ElementAt(i).Value; }
        }

        public RibbonGroup this[string name]
        {
            get { return _groups[name]; }
        }

        #endregion

        #region Costruttori

        internal GroupsCollection(ToolsExcelRibbon ribbon)
        {
            _ribbon = ribbon;
        }

        #endregion

        #region Metodi

        public void Add(RibbonGroup group)
        {
            _groups.Add(group.Name, group);
        }

        public IEnumerator GetEnumerator()
        {
            return new GroupEnumerator(_ribbon);
        }

        public IEnumerable<KeyValuePair<string, RibbonGroup>> AsEnumerable()
        {
            return _groups.AsEnumerable();
        }

        #endregion
    }
    public class GroupEnumerator : IEnumerator
    {
        #region Variabili

        private ToolsExcelRibbon _ribbon;
        private int _pos = -1;
        private int _max = -1;

        #endregion

        #region Costruttori

        public GroupEnumerator(ToolsExcelRibbon ribbon)
        {
            _ribbon = ribbon;
            _max = ribbon.Groups.Count;
        }

        #endregion

        #region Metodi

        public object Current
        {
            get { return _ribbon.Groups[_pos]; }
        }
        public bool MoveNext()
        {
            _pos++;
            return _pos < _max;
        }
        public void Reset()
        {
            _pos = -1;
        }

        #endregion
    }

    #endregion
}
