using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Data;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Iren.ToolsExcel.Core;
using System.Text.RegularExpressions;
using System.Configuration;
using Word = Microsoft.Office.Interop.Word;
using DataTable = System.Data.DataTable;
using DataRow = System.Data.DataRow;
using DataView = System.Data.DataView;
using Iren.RiMoST.Properties;
using System.Drawing;
using System.Deployment.Application;

namespace Iren.RiMoST
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        #region Variabili

        private Office.IRibbonUI ribbon;
        FormAnnullaModifica _formAnnullaModifica;
        internal int _cbAnniDispCount = 0;
        internal int _cbAnniDispIndex = 0;
        internal List<string> _cbAnniDispLabels;
        internal string _cbAnniDispValue = "";
        internal System.Version _appV;
        internal System.Version _coreV;
        internal bool _chkIsDraftEnabled = true;
        internal bool _chkIsDraft = false;
        internal bool _btnSalvaBozzaEnabled = true;
        internal bool _btnRefreshEnabled = true;       

        #endregion

        #region Costruttori

        public Ribbon()
        {
        }

        #endregion

        #region Metodi Privati

        private System.Version getCurrentV()
        {
            try
            {
                return ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
            catch (Exception)
            {
                return Assembly.GetExecutingAssembly().GetName().Version;
            }
        }

        private bool EmptyFields()
        {
            if (Globals.ThisDocument.txtOggetto.Text == "" || Globals.ThisDocument.txtDescrizione.Text == "")
            {
                MessageBox.Show("Alcuni campi obbligatori non sono stati compilati. Compilare i campi evidenziati!", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Error);

                Globals.ThisDocument.Application.ScreenUpdating = false;

                ThisDocument.ToNormal("Oggetto", Word.WdColorIndex.wdBlack, "*");
                ThisDocument.ToNormal("Descrizione", Word.WdColorIndex.wdBlack, "*");

                if (Globals.ThisDocument.txtOggetto.Text == "")
                    ThisDocument.Highlight("Oggetto", Word.WdColorIndex.wdRed, "*");

                if (Globals.ThisDocument.txtDescrizione.Text == "")
                    ThisDocument.Highlight("Descrizione", Word.WdColorIndex.wdRed, "*");

                Globals.ThisDocument.Application.ScreenUpdating = true;

                return true;
            }

            Globals.ThisDocument.Application.ScreenUpdating = false;

            ThisDocument.ToNormal("Oggetto", Word.WdColorIndex.wdBlack, "*");
            ThisDocument.ToNormal("Descrizione", Word.WdColorIndex.wdBlack, "*");

            Globals.ThisDocument.Application.ScreenUpdating = true;

            return false;
        }

        private void getAvailableID()
        {
            if (ThisDocument._db.OpenConnection())
            {
                DataTable dt = ThisDocument._db.Select("spGetFirstAvailableID", "@IdStruttura=" + ThisDocument._idStruttura);
                Globals.ThisDocument.lbIdRichiesta.LockContents = false;
                Globals.ThisDocument.lbIdRichiesta.Text = dt.Rows[0][0].ToString();
                Globals.ThisDocument.lbIdRichiesta.LockContents = true;
                ThisDocument._db.CloseConnection();
            }
        }

        private void Print()
        {
            object missing = Missing.Value;

            if (Globals.ThisDocument.Application.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFilePrint].Show() == 1)
            {
                Globals.ThisDocument.PrintOut(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            }
        }

        private void ChangeBozzaVisibility(bool visible)
        {
            //if (!visible)
            //{
            //    Globals.ThisDocument.lbBozza.Text = "";
            //    Globals.ThisDocument.lbBozza.Image = null;
            //}
            //else
            //{
            //    Globals.ThisDocument.lbBozza.Text = "Bozza";
            //    Globals.ThisDocument.lbBozza.Image = Resources.Editing_Edit_icon;
            //}
        }

        #endregion

        #region Membri IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("RiMoST2.Ribbon.xml");
        }

        #endregion

        #region Callback della barra multifunzione

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            if(ThisDocument._db.OpenConnection())
            {
                DataTable dt = ThisDocument._db.Select("spGetAvailableYears", "@IdStruttura=" + ThisDocument._idStruttura);
                _cbAnniDispLabels = new List<string>();
                foreach (DataRow r in dt.Rows)
                {
                    RibbonDropDownItem i = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    i.Label = r["Anno"].ToString();
                    _cbAnniDispLabels.Add(r["Anno"].ToString());
                }
                _cbAnniDispCount = _cbAnniDispLabels.Count;
                _appV = getCurrentV();
                _coreV = ThisDocument._db.GetCurrentV();
                ThisDocument._db.CloseConnection();
            }
        }

        public void btnReset_Click(Office.IRibbonControl control)
        {
            if (MessageBox.Show("Sicuro di voler cancellare il contenuto dei campi?", "Cancellare campi?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                Globals.ThisDocument.cmbStrumento.DropDownListEntries[0].Select();                
                Globals.ThisDocument.cmbStrumento.LockContents = false;
                Globals.ThisDocument.txtDescrizione.Text = "";
                Globals.ThisDocument.txtOggetto.Text = "";
                Globals.ThisDocument.txtNote.Text = "";

                _btnRefreshEnabled = true;
                _chkIsDraft = false;
                _chkIsDraftEnabled = true;
                this.ribbon.InvalidateControl("chkIsDraft");
                this.ribbon.InvalidateControl("btnSalvaBozza");
                this.ribbon.InvalidateControl("btnRefresh");
                getAvailableID();
            }
        }
        public void btnInvia_Click(Office.IRibbonControl control)
        {
            object oTrue = true;
            object oFalse = false;
            object missing = Missing.Value;
            if (ThisDocument._db.OpenConnection())
            {
                DataView dv = ThisDocument._db.Select("spGetRichiesta", "@IdRichiesta=" + Globals.ThisDocument.lbIdRichiesta.Text + ";@IdStruttura=" + ThisDocument._idStruttura).DefaultView;
                dv.RowFilter = "IdTipologiaStato <> 7";
                if (dv.Count > 0)
                {
                    MessageBox.Show("Esiste già una richiesta con lo stesso codice. Premere sul tasto di refresh per ottenerne uno nuovo", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (!EmptyFields())
                    {
                        if (_chkIsDraft || MessageBox.Show("Sicuro di voler inviare il documento?", "Stampa e invia?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                        {
                            Globals.ThisDocument.Application.ScreenUpdating = false;

                            ThisDocument.ToNormal("Oggetto", Word.WdColorIndex.wdBlack, "*");
                            ThisDocument.ToNormal("Descrizione", Word.WdColorIndex.wdBlack, "*");

                            Globals.ThisDocument.Application.ScreenUpdating = true;

                            _btnRefreshEnabled = !_chkIsDraft;
                            if (!_chkIsDraft)
                                _chkIsDraftEnabled = false;
                            this.ribbon.InvalidateControl("chkIsDraft");
                            this.ribbon.InvalidateControl("btnRefresh");

                            string saveName = ConfigurationManager.AppSettings["saveNameFormat"];
                            foreach (Match m in Regex.Matches(saveName, @"(\[[^\[\]]*\])"))
                            {
                                try
                                {
                                    Control c = (Control)Globals.ThisDocument.Controls[m.Value.Replace("[", "").Replace("]", "")];
                                    saveName = saveName.Replace(m.Value, c.Text);
                                }
                                catch (ArgumentOutOfRangeException)
                                {

                                }
                            }
                            string name = Regex.Replace(saveName, @"([^\.\-_a-zA-Z0-9]+)", "_");
                            object savePath = Path.Combine(ConfigurationManager.AppSettings["savePath"], name + ".pdf");
                            object format = Word.WdSaveFormat.wdFormatPDF;
                            try
                            {
                                if (!_chkIsDraft)
                                    Globals.ThisDocument.SaveAs2(ref savePath, ref format, ref oTrue, ref missing, ref oFalse,
                                        ref missing, ref oFalse, ref missing, ref missing, ref oFalse, ref oFalse, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing);

                                DateTime dataInvio = DateTime.Parse(Globals.ThisDocument.lbDataInvio.Text);
                                string idApplicazione = Globals.ThisDocument.cmbStrumento.DropDownListEntries[0].Value;

                                QryParams parameters = new QryParams()
                            {
                                {"@IdRichiesta", Globals.ThisDocument.lbIdRichiesta.Text},
                                {"@IdStruttura", ThisDocument._idStruttura},
                                {"@DataInvio", dataInvio.ToString("yyyyMMdd")},
                                {"@IdTipologiaStato", _chkIsDraft ? 7:1},
                                {"@IdApplicazione", idApplicazione},
                                {"@Oggetto", Globals.ThisDocument.txtOggetto.Text.Trim()},
                                {"@Descr", Globals.ThisDocument.txtDescrizione.Text.Trim()},
                                {"@Note", Globals.ThisDocument.txtNote.Text.Trim()},
                                {"@NomeFile", savePath}
                            };

                                ThisDocument._db.Insert("spSaveRichiestaModifica", parameters);

                                if (!_chkIsDraft)
                                    Print();
                            }
                            catch (Exception)
                            {
                                MessageBox.Show("Salvataggio non riuscito... Riprovare più tardi.", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
                ThisDocument._db.CloseConnection();
            }
            else
            {
                MessageBox.Show("Salvataggio non riuscito... Riprovare più tardi.", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
        public void btnChiudi_Click(Office.IRibbonControl control)
        {
            Globals.ThisDocument.CloseWithoutSaving();
        }
        public void btnRefresh_Click(Office.IRibbonControl control)
        {
            getAvailableID();
        }
        public void btnPrint_Click(Office.IRibbonControl control)
        {
            Print();
        }
        public void btnAnnulla_Click(Office.IRibbonControl control)
        {
            if (_formAnnullaModifica == null || _formAnnullaModifica.IsDisposed)
            {
                _formAnnullaModifica = new FormAnnullaModifica(_cbAnniDispValue);
                _formAnnullaModifica.Show();
            }
            _formAnnullaModifica.WindowState = FormWindowState.Normal;
            _formAnnullaModifica.Focus();
        }
        public void btnModifica_Click(Office.IRibbonControl control)
        {
            SelezionaModifica selMod = new SelezionaModifica(_cbAnniDispValue, _chkIsDraft, _btnRefreshEnabled);
            selMod.ShowDialog();
            _chkIsDraft = selMod._chkIsDraft;
            _btnRefreshEnabled = selMod._btnRefreshEnabled;
            this.ribbon.InvalidateControl("btnRefresh");
            this.ribbon.InvalidateControl("chkIsDraft");
            selMod.Dispose();
        }
        public void chkIsDraft_Click(Office.IRibbonControl control, bool pressed)
        {
            ChangeBozzaVisibility(pressed);
            _chkIsDraft = pressed;
            _btnRefreshEnabled = !pressed;
            this.ribbon.InvalidateControl("btnRefresh");
        }
        
        public bool chkIsDraft_getPressed(Office.IRibbonControl control) 
        {
            ChangeBozzaVisibility(_chkIsDraft);
            return _chkIsDraft;
        }
        
        public int cbAnniDisp_ItemCount(Office.IRibbonControl control)
        {
            return _cbAnniDispCount;
        }
        public string cbAnniDisp_ItemLabel(Office.IRibbonControl control, int i)
        {
            return _cbAnniDispLabels[i];
        }
        public int cbAnniDisp_getSelectedItemIndex(Office.IRibbonControl control)
        {
            return _cbAnniDispIndex;
        }
        public void cbAnniDisp_onAction(Office.IRibbonControl control, string itemID, int itemIndex)
        {
            _cbAnniDispValue = _cbAnniDispLabels[itemIndex];
            _cbAnniDispIndex = itemIndex;
        }
        public string lbVersioneApp_getLabel(Office.IRibbonControl control)
        {
            return "  App v" + _appV.ToString();
        }
        public string lbCoreV_getLabel(Office.IRibbonControl control)
        {
            return "  Core v" + _coreV.ToString();
        }
        public bool chkIsDraft_getEnabled(Office.IRibbonControl control)
        {
            return _chkIsDraftEnabled;
        }
        public bool btnRefresh_getEnabled(Office.IRibbonControl control)
        {
            return _btnRefreshEnabled;
        }

        public Bitmap btnReset_getImage(Office.IRibbonControl control)
        {
            return Resources.Reset_icon;
        }
        public Bitmap btnInvia_getImage(Office.IRibbonControl control)
        {
            return Resources.Send_icon;
        }
        public Bitmap chkIsDraft_getImage(Office.IRibbonControl control)
        {
            return Resources.draft_icon;
        }
        public Bitmap btnChiudi_getImage(Office.IRibbonControl control)
        {
            return Resources.Close_icon;
        }
        public Bitmap btnRefresh_getImage(Office.IRibbonControl control)
        {
            return Resources.Refresh_icon;
        }
        public Bitmap btnPrint_getImage(Office.IRibbonControl control)
        {
            return Resources.Print_icon;
        }
        public Bitmap btnAnnulla_getImage(Office.IRibbonControl control)
        {
            return Resources.Bin_icon;
        }
        public Bitmap btnModifica_getImage(Office.IRibbonControl control)
        {
            return Resources.edit_icon;
        }
        public Bitmap cbAnniDisponibili_getImage(Office.IRibbonControl control)
        {
            return Resources.calendar_icon;
        }

        #endregion

        #region Supporti

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
