using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Data;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Deployment.Application;
using System.Configuration;
using System.Text.RegularExpressions;
using System.IO;
using Iren.PSO.Core;

namespace Iren.RiMoST
{
    public partial class RiMoSTRibbon
    {
        #region Variabili 

        private bool _isBozza;
        private Dictionary<string, string> _originalControlText = new Dictionary<string,string>();

        #endregion

        #region Proprietà

        public bool isBozza
        {
            get { return _isBozza; }
            set
            {
                _isBozza = value;
                chkIsDraft.Checked = _isBozza;

                Globals.ThisDocument.headerIdRichiesta.LockContents = false;
                Globals.ThisDocument.headerIdRichiesta.Text = (_isBozza ? "Bozza - " : "") + Globals.ThisDocument.lbIdRichiesta.Text;
                Globals.ThisDocument.headerIdRichiesta.LockContents = true;

                Globals.ThisDocument.lbBozza.LockContents = false;
                Globals.ThisDocument.lbBozza.Range.Font.Hidden = _isBozza ? 0 : 1;
                Globals.ThisDocument.lbBozza.LockContents = true;
            }
        }

        #endregion

        #region Metodi

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

                ThisDocument.ToNormal(Globals.ThisDocument.lbOggetto);
                ThisDocument.ToNormal(Globals.ThisDocument.lbDescrizione);

                if (Globals.ThisDocument.txtOggetto.Text == "")
                    ThisDocument.Highlight(Globals.ThisDocument.lbOggetto);

                if (Globals.ThisDocument.txtDescrizione.Text == "")
                    ThisDocument.Highlight(Globals.ThisDocument.lbDescrizione);

                Globals.ThisDocument.Application.ScreenUpdating = true;

                return true;
            }

            Globals.ThisDocument.Application.ScreenUpdating = false;

            ThisDocument.ToNormal(Globals.ThisDocument.lbOggetto);
            ThisDocument.ToNormal(Globals.ThisDocument.lbDescrizione);

            Globals.ThisDocument.Application.ScreenUpdating = true;

            return false;
        }

        public void getAvailableID()
        {
            if (ThisDocument.DB.OpenConnection())
            {
                DataTable dt = ThisDocument.DB.Select("spGetFirstAvailableID", "@IdStruttura=" + ThisDocument._idStruttura) ?? new DataTable();
                
                Globals.ThisDocument.lbIdRichiesta.LockContents = false;
                Globals.ThisDocument.lbIdRichiesta.Text = dt.Rows[0][0].ToString();
                Globals.ThisDocument.lbIdRichiesta.LockContents = true;
                
                Globals.ThisDocument.headerIdRichiesta.LockContents = false;
                Globals.ThisDocument.headerIdRichiesta.Text = dt.Rows[0][0].ToString();
                Globals.ThisDocument.headerIdRichiesta.LockContents = true;
                
                ThisDocument.DB.CloseConnection();
            }
        }

        private bool CutContent(Microsoft.Office.Tools.Word.RichTextContentControl ctrl)
        {
            int length = 0;

            if (ctrl.Equals(Globals.ThisDocument.txtDescrizione))
                length = ThisDocument.DESCRIZIONE_MAX_LEN;
            else if (ctrl.Equals(Globals.ThisDocument.txtNote))
                length = ThisDocument.NOTE_MAX_LEN;
            else if (ctrl.Equals(Globals.ThisDocument.txtOggetto))
                length = ThisDocument.OGGETTO_MAX_LEN;

            if (ctrl.Text.Length > length)
            {
                if (!_originalControlText.ContainsKey(ctrl.ID))
                    _originalControlText.Add(ctrl.ID, ctrl.Text);
                else
                    _originalControlText[ctrl.ID] = ctrl.Text;

                ctrl.Text = ctrl.Text.Substring(0, Math.Min(ctrl.Text.Length, length));

                return true;
            }
            return false;
        }
        private void UncutContent(Microsoft.Office.Tools.Word.RichTextContentControl ctrl)
        {
            ctrl.Text = _originalControlText[ctrl.ID];

            int length = 0;
            if (ctrl.Equals(Globals.ThisDocument.txtDescrizione))
                length = ThisDocument.DESCRIZIONE_MAX_LEN;
            else if (ctrl.Equals(Globals.ThisDocument.txtNote))
                length = ThisDocument.NOTE_MAX_LEN;
            else if (ctrl.Equals(Globals.ThisDocument.txtOggetto))
                length = ThisDocument.OGGETTO_MAX_LEN;

            Globals.ThisDocument.FormatTextOverDimension(ctrl, length, ctrl.Text.Length > length);
        }

        private void Print()
        {
            object missing = Missing.Value;
            bool cut1 = false;
            bool cut2 = false;
            bool cut3 = false;

            Globals.ThisDocument.ThisApplication.ScreenUpdating = false;

            if (Globals.ThisDocument.txtNote.ShowingPlaceholderText)
                Globals.ThisDocument.txtNote.Range.Font.Hidden = 1;
            else
                cut1 = CutContent(Globals.ThisDocument.txtNote);

            if (Globals.ThisDocument.txtOggetto.ShowingPlaceholderText)
                Globals.ThisDocument.txtOggetto.Range.Font.Hidden = 1;
            else
                cut2 = CutContent(Globals.ThisDocument.txtOggetto);

            if (Globals.ThisDocument.txtDescrizione.ShowingPlaceholderText)
                Globals.ThisDocument.txtDescrizione.Range.Font.Hidden = 1;
            else
                cut3 = CutContent(Globals.ThisDocument.txtDescrizione);


            if (Globals.ThisDocument.Application.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFilePrint].Show() == 1)
            {
                Globals.ThisDocument.PrintOut(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            }

            if (cut1) UncutContent(Globals.ThisDocument.txtNote);
            if (cut2) UncutContent(Globals.ThisDocument.txtOggetto);
            if (cut3) UncutContent(Globals.ThisDocument.txtDescrizione);

            Globals.ThisDocument.txtNote.Range.Font.Hidden = 0;
            Globals.ThisDocument.txtOggetto.Range.Font.Hidden = 0;
            Globals.ThisDocument.txtDescrizione.Range.Font.Hidden = 0;

            Globals.ThisDocument.ThisApplication.ScreenUpdating = true;
        }

        #endregion

        private void RiMoSTRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            if (ThisDocument.DB.OpenConnection())
            {
                DataTable dt = ThisDocument.DB.Select("spGetAvailableYears", "@IdStruttura=" + ThisDocument._idStruttura) ?? new DataTable();
                foreach (DataRow r in dt.Rows)
                {
                    RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    item.Label = r["Anno"].ToString();
                    Globals.Ribbons.RiMoSTRibbon.cbAnniDisponibili.Items.Add(item);
                }
                lbCoreV.Label = "  Core v" + ThisDocument.DB.GetCurrentV().ToString();
                lbVersioneApp.Label = "  App v" + getCurrentV().ToString();

                isBozza = true;

                ThisDocument.DB.CloseConnection();
            }
        }
        private void btnReset_Click(object sender, RibbonControlEventArgs e)
        {
            if (MessageBox.Show("Sicuro di voler cancellare il contenuto dei campi?", "Cancellare campi?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                Globals.ThisDocument.dropDownStrumenti.LockContents = false;
                Globals.ThisDocument.dropDownStrumenti.DropDownListEntries[1].Select();

                Globals.ThisDocument.txtDescrizione.Text = "";
                Globals.ThisDocument.txtOggetto.Text = "";
                Globals.ThisDocument.txtNote.Text = "";

                chkIsDraft.Enabled = true;

                getAvailableID();
                isBozza = true;
            }
        }
        private void btnInvia_Click(object sender, RibbonControlEventArgs e)
        {
            object oTrue = true;
            object oFalse = false;
            object missing = Missing.Value;

            if (ThisDocument.DB.OpenConnection())
            {
                DataView dv = (ThisDocument.DB.Select("spGetRichiesta", "@IdRichiesta=" + Globals.ThisDocument.lbIdRichiesta.Text + ";@IdStruttura=" + ThisDocument._idStruttura) ?? new DataTable()).DefaultView;
                dv.RowFilter = "IdTipologiaStato <> 7";
                if (dv.Count > 0)
                {
                    MessageBox.Show("Esiste già una richiesta con lo stesso codice. Premere sul tasto di refresh per ottenerne uno nuovo", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (!EmptyFields())
                    {
                        if (isBozza || MessageBox.Show("Sicuro di voler inviare il documento?", "Stampa e invia?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                        {
                            Globals.ThisDocument.Application.ScreenUpdating = false;

                            ThisDocument.ToNormal(Globals.ThisDocument.lbOggetto);
                            ThisDocument.ToNormal(Globals.ThisDocument.lbDescrizione);

                            Globals.ThisDocument.Application.ScreenUpdating = true;

                            btnRefresh.Enabled = !isBozza;
                            if (!isBozza)
                                chkIsDraft.Enabled = false;

                            string saveName = ConfigurationManager.AppSettings["saveNameFormat"];
                            foreach (Match m in Regex.Matches(saveName, @"(\[[^\[\]]*\])"))
                            {
                                try
                                {
                                    Microsoft.Office.Tools.Word.RichTextContentControl c =
                                        (Microsoft.Office.Tools.Word.RichTextContentControl)Globals.ThisDocument.Controls[m.Value.Replace("[", "").Replace("]", "")];
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
                                if (!isBozza)
                                    Globals.ThisDocument.SaveAs2(ref savePath, ref format, ref oTrue, ref missing, ref oFalse,
                                        ref missing, ref oFalse, ref missing, ref missing, ref oFalse, ref oFalse, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing);

                                DateTime dataInvio = DateTime.Parse(Globals.ThisDocument.lbDataInvio.Text);
                                //string idApplicazione = Globals.ThisDocument.dropDownStrumenti.DropDownListEntries[1].Value;
                                string idApplicazione = Globals.ThisDocument.dropDownStrumenti.DropDownListEntries.OfType<Microsoft.Office.Interop.Word.ContentControlListEntry>().First(c => c.Text == Globals.ThisDocument.dropDownStrumenti.Text).Value;

                                ThisDocument.DB.Insert("spSaveRichiestaModifica", new QryParams()
                                    {
                                        {"@IdRichiesta", Globals.ThisDocument.lbIdRichiesta.Text},
                                        {"@IdStruttura", ThisDocument._idStruttura},
                                        {"@DataInvio", dataInvio.ToString("yyyyMMdd")},
                                        {"@IdTipologiaStato", isBozza ? 7:1},
                                        {"@IdApplicazione", idApplicazione},
                                        {"@Oggetto", Globals.ThisDocument.txtOggetto.Text.Trim()},
                                        {"@Descr", Globals.ThisDocument.txtDescrizione.Text.Trim()},
                                        {"@Note", Globals.ThisDocument.txtNote.Text.Trim()},
                                        {"@NomeFile", savePath}
                                    });

                                if(!isBozza)
                                    MessageBox.Show("Richiesta presa in carico!", "Perfetto!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                else
                                    MessageBox.Show("Bozza salvata!", "Perfetto!!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                if (!isBozza)
                                    Print();
                            }
                            catch (Exception)
                            {
                                MessageBox.Show("Salvataggio non riuscito... Riprovare più tardi.", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
                ThisDocument.DB.CloseConnection();
            }
            else
            {
                MessageBox.Show("Errore nella connessione al DB... Riprovare più tardi.", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
        private void btnChiudi_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisDocument.CloseWithoutSaving();
        }
        private void btnRefresh_Click(object sender, RibbonControlEventArgs e)
        {
            getAvailableID();
            chkIsDraft.Enabled = true;
            isBozza = true;
        }
        private void btnPrint_Click(object sender, RibbonControlEventArgs e)
        {
            Print();
        }
        private void btnAnnulla_Click(object sender, RibbonControlEventArgs e)
        {
            FormAnnullaModifica formAnnulla = new FormAnnullaModifica(cbAnniDisponibili.Text);
            formAnnulla.ShowDialog();
            formAnnulla.Dispose();
        }
        private void btnModifica_Click(object sender, RibbonControlEventArgs e)
        {
            SelezionaModifica selMod = new SelezionaModifica(cbAnniDisponibili.Text, isBozza, btnRefresh.Enabled);
            selMod.ShowDialog();
            isBozza = selMod._chkIsDraft;
            btnRefresh.Enabled = selMod._btnRefreshEnabled;
            selMod.Dispose();
        }
        private void chkIsDraft_Click(object sender, RibbonControlEventArgs e)
        {
            isBozza = !isBozza;
        }
    }
}
