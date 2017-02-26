using Iren.PSO.Base;
using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Funzione di esportazione custom.
    /// </summary>
    class Esporta : AEsporta
    {
        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif, string[] mercati)
        {
            DataView entitaAzione = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            if (entitaAzione.Count == 0)
                return false;

            switch (siglaAzione.ToString())
            {
                case "MAIL":
                    Workbook.ScreenUpdating = false;
                    DefinedNames mainDefinedNames = new DefinedNames("Main");
                    //TODO verificare se è sempre aggiornato
                    //unico caso che non aggiorna è se carico e faccio invia mail conseguentemente

                    Aggiorna a = new Aggiorna();
                    a.AggiornaPrevisioneRiepilogo();
                
                    //salvo i dati 
                    Riepilogo r = new Riepilogo();
                    r.SalvaPrevisione();

                    if (InviaMail(mainDefinedNames, siglaEntita))
                    {

                    }

                    Workbook.ScreenUpdating = true;
                    break;
            }
            return true;
        }

        protected bool InviaMail(DefinedNames definedNames, object siglaEntita) 
        {
            string fileNameFull = "";
            string fileName = "";
            try
            {
                fileName = @"PrevisioneGAS_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                fileNameFull = Environment.ExpandEnvironmentVariables(@"%TEMP%\" + fileName);

                Excel.Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Add();

                Workbook.Main.Range[Range.GetRange(definedNames.GetFirstRow(), definedNames.GetFirstCol(), definedNames.GetRowOffset(), definedNames.GetColOffsetRiepilogo()).ToString()].Copy();
                wb.Sheets[1].Range["B2"].PasteSpecial();

                wb.Sheets[1].UsedRange.ColumnWidth = 17;
                wb.Sheets[1].Range["A1"].Select();
                wb.SaveAs(fileNameFull, Excel.XlFileFormat.xlExcel8);
                wb.Close();
                Marshal.ReleaseComObject(wb);

                var config = Workbook.GetUsrConfigElement("destMailTest");
                string mailTo = config.Test;
                string mailCC = "";

                DataView entitaProprieta = new DataView(Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA]);

                if (Workbook.Ambiente == Simboli.PROD)
                {
                    entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'PREV_CONSUMO_GAS_MAIL_TO' AND IdApplicazione = " + Workbook.IdApplicazione;
                    
                    if(entitaProprieta.Count > 0)
                        mailTo = entitaProprieta[0]["Valore"].ToString();
                    
                    entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'PREV_CONSUMO_GAS_MAIL_CC' AND IdApplicazione = " + Workbook.IdApplicazione;
                    
                    if(entitaProprieta.Count > 0)
                        mailCC = entitaProprieta[0]["Valore"].ToString();
                }
                if (DataBase.OpenConnection())
                {
                    Outlook.Application outlook = GetOutlookInstance();
                    Outlook._MailItem mail = outlook.CreateItem(Outlook.OlItemType.olMailItem);

                    config = Workbook.GetUsrConfigElement("oggettoMail");
                    string oggetto = config.Value.Replace("%DATA%", DateTime.Now.ToString("dd-MM-yyyy")).Replace("%ORA%", DateTime.Now.ToString("HH:mm"));
                    config = Workbook.GetUsrConfigElement("messaggioMail");
                    string messaggio = config.Value.Replace("%NOMEUTENTE%", Workbook.NomeUtente);
                    messaggio = Regex.Replace(messaggio, @"^[^\S\r\n]+", "", RegexOptions.Multiline);


                    ////TODO check se manda sempre con lo stesso account...
                    Outlook.Account senderAccount = outlook.Session.Accounts[1];
                    foreach (Outlook.Account account in outlook.Session.Accounts)
                    {
                        if (account.DisplayName == "Bidding")
                            senderAccount = account;
                    }
                    mail.SendUsingAccount = senderAccount;
                    mail.Subject = oggetto;
                    mail.Body = messaggio;
                    foreach (string dest in mailTo.Split(';'))
                        if (dest.Trim() != "")
                            mail.Recipients.Add(dest.Trim());
                    mail.CC = mailCC;
                    mail.Attachments.Add(fileNameFull);

                    mail.Send();

                    File.Delete(fileNameFull);
                }
                else
                {
                    string emailFolder = @"C:\Emergenza\Email\" + Simboli.NomeApplicazione;

                    if (!Directory.Exists(emailFolder))
                        Directory.CreateDirectory(emailFolder);

                    File.Move(fileNameFull, Path.Combine(emailFolder, fileName));
                }
            }
            catch(Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "PrevisioneGAS.Esporta.InvioMail: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                if(File.Exists(fileName))
                    File.Delete(fileName);

                return false;
            }

            return true;
        }
    }
}
