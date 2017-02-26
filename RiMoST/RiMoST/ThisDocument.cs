using System;
using System.Configuration;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Specialized;
using DataTable = System.Data.DataTable;
using DataRow = System.Data.DataRow;
using DataView = System.Data.DataView;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using Iren.PSO.Core;
using System.IO;
using Microsoft.Office.Tools.Ribbon;

namespace Iren.RiMoST
{
    public partial class ThisDocument
    {
        #region Variabili

        private static DataBase _db;
        public static string _idStruttura;
        private DataTable _dtApplicazioni;
        public const int OGGETTO_MAX_LEN = 255;
        public const int DESCRIZIONE_MAX_LEN = 4000;
        public const int NOTE_MAX_LEN = 255;
        
        private Microsoft.Office.Tools.Word.GroupContentControl groupControl1;

        #endregion

        #region Proprietà

        public DataTable Applicazioni { get { return _dtApplicazioni; } }
        public static DataBase DB { get { return _db; } }

        #endregion

        #region Callbacks

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            SetAppConfigEnvironment();

            _idStruttura = ConfigurationManager.AppSettings["idStruttura"];
            string[] users = new string[1];
            if (ConfigurationManager.AppSettings["utentiVisto"] != null)
                users = ConfigurationManager.AppSettings["utentiVisto"].Split(',');

            int rowNum = (int)Math.Ceiling(users.Length / 5.0);
            int colNum = (int)Math.Ceiling((decimal)users.Length / rowNum);

            Word.Table tb = this.Tables[1];
            
            tb.Rows[tb.Rows.Count].Cells.Split(rowNum, colNum);
            int i = 0, j = 0;
            foreach (string usr in users)
            {
                tb.Cell((tb.Rows.Count - rowNum) + i + 1, (j % colNum) + 1).Range.Text = usr;

                if ((++j % colNum) == 0)
                    i++;
            }
            

            _db = new DataBase(ConfigurationManager.AppSettings["DB"], false);
            if(_db.OpenConnection())
            {
                _dtApplicazioni = _db.Select("spGetApplicazioniDisponibili", "@IdStruttura=" + _idStruttura) ?? new DataTable();
                foreach (DataRow r in _dtApplicazioni.Rows)
                {
                    dropDownStrumenti.DropDownListEntries.Add(r["DesApplicazione"].ToString(), r["IdApplicazione"].ToString());
                }

                dropDownStrumenti.DropDownListEntries[1].Select();

                Globals.Ribbons.RiMoSTRibbon.getAvailableID();                

                lbDataInvio.LockContents = false;
                lbDataInvio.Text = DateTime.Now.ToShortDateString();
                lbDataInvio.LockContents = true;

                object what = Word.WdGoToItem.wdGoToLine;
                object which = Word.WdGoToDirection.wdGoToLast;
                object missing = Missing.Value;

                _db.CloseConnection();
            }

            //disabilito la modifica per gli elementi strutturali del documento
            tb.Select();
            groupControl1 = this.Controls.AddGroupContentControl("TabellaStruttura");

            dropDownStrumenti.Range.Select();
            this.ThisApplication.Options.PasteFormatWithinDocument = WdPasteOptions.wdKeepTextOnly;
        }
        private void ThisDocument_BeforeClose(object sender, System.ComponentModel.CancelEventArgs e)
        {
            CloseWithoutSaving();
        }
        #pragma warning disable 0467
        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
            ThisApplication.Quit();
        }
        #pragma warning restore 0467
        private void RichText_Exiting(object sender, ContentControlExitingEventArgs e)
        {
            Microsoft.Office.Tools.Word.RichTextContentControl ctrl = (Microsoft.Office.Tools.Word.RichTextContentControl)sender;

            string messaggio = "";
            bool overDimension = false;
            int length = 0;
            if (ctrl.Equals(txtOggetto) && ctrl.Text.Length > OGGETTO_MAX_LEN)
            {
                messaggio = "La lunghezza dell'oggetto supera i caratteri consentiti. Se si procede così il testo salvato risulterà parziale.";
                overDimension = true;
                length = OGGETTO_MAX_LEN;
            }
            else if (ctrl.Equals(txtDescrizione) && ctrl.Text.Length > DESCRIZIONE_MAX_LEN)
            {
                messaggio = "La lunghezza della descrizione supera i caratteri consentiti. Se si procede così il testo salvato risulterà parziale.";
                overDimension = true;
                length = DESCRIZIONE_MAX_LEN;
            }
            else if (ctrl.Equals(txtNote) && ctrl.Text.Length > NOTE_MAX_LEN)
            {
                messaggio = "La lunghezza delle note supera i caratteri consentiti. Se si procede così il testo salvato risulterà parziale.";
                overDimension = true;
                length = NOTE_MAX_LEN;
            }

            FormatTextOverDimension(ctrl, length, overDimension);
            if (overDimension)
                System.Windows.Forms.MessageBox.Show(messaggio, "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        #endregion

        #region Metodi

        public void FormatTextOverDimension(Microsoft.Office.Tools.Word.RichTextContentControl ctrl, int length, bool overDimension)
        {
            if (overDimension)
            {
                string text = ctrl.Text;
                Word.Range rng = ctrl.Range;
                rng.Text = text.Substring(0, length);
                rng.Font.ColorIndex = WdColorIndex.wdBlack;
                rng.Collapse(WdCollapseDirection.wdCollapseEnd);
                rng.Text = text.Substring(length, text.Length - length);
                rng.Font.ColorIndex = WdColorIndex.wdRed;
            }
            else
            {
                if(!ctrl.ShowingPlaceholderText)
                    ctrl.Range.Font.ColorIndex = WdColorIndex.wdAuto;
            }
        }
        private void SetAppConfigEnvironment()
        {
            string file = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RiMoST/RiMoST.config");
            if (File.Exists(AppDomain.CurrentDomain.GetData("APP_CONFIG_FILE").ToString()))
            {
                Directory.CreateDirectory(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RiMoST"));
                File.Copy(AppDomain.CurrentDomain.GetData("APP_CONFIG_FILE").ToString(), file, true);
                File.Delete(AppDomain.CurrentDomain.GetData("APP_CONFIG_FILE").ToString());
            }
            AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", file);
            //CryptHelper.CryptSection("connectionStrings");
        }
        public void CloseWithoutSaving()
        {
            object saveMod = WdSaveOptions.wdDoNotSaveChanges;
            object missing = Missing.Value;
            this.Close(ref saveMod, ref missing, ref missing);
        }
        public static void Highlight(Microsoft.Office.Tools.Word.RichTextContentControl ctrl)
        {
            ctrl.LockContents = false;
            ctrl.Text = ctrl.Text + "*";
            ctrl.Range.Font.ColorIndex = WdColorIndex.wdRed;
            ctrl.LockContents = true;
        }
        public static void ToNormal(Microsoft.Office.Tools.Word.RichTextContentControl ctrl)
        {
            ctrl.LockContents = false;
            ctrl.Text = ctrl.Text.Replace("*","");
            ctrl.Range.Font.ColorIndex = WdColorIndex.wdBlack;
            ctrl.LockContents = true;
        }

        #endregion

        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.txtDescrizione.Exiting += new Microsoft.Office.Tools.Word.ContentControlExitingEventHandler(this.RichText_Exiting);
            this.txtOggetto.Exiting += new Microsoft.Office.Tools.Word.ContentControlExitingEventHandler(this.RichText_Exiting);
            this.txtNote.Exiting += new Microsoft.Office.Tools.Word.ContentControlExitingEventHandler(this.RichText_Exiting);
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);
            this.BeforeClose += new System.ComponentModel.CancelEventHandler(this.ThisDocument_BeforeClose);

        }

        #endregion
    }
}
