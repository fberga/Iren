using Iren.PSO.Base;
using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Classe base con i metodi per la creazione di un foglio contenente dati riferiti a impianti.
    /// </summary>
    public class Sheet : Base.Sheet
    {
        #region Costruttori

        public Sheet(Excel.Worksheet ws)
            : base(ws)
        {

        }
        #endregion

        #region Metodi

        //06/02/2017 MOD: nascondo le righe dei mercati non di competenza.
        public void HideMarketRows()
        {
            /* Recupero mercato attivo al momento:
             *  - Prendo il primo mercato disponibile con chiusura > di ora
             */
            int hour = DateTime.Now.Hour;

            //08/02/2017 FIX: messa logica di selezione mercato univoca
            //string mercatoAttivo = Simboli.GetActiveMarket(hour);
            //09/02/2017 MOD: gestione manuale del mercato
            string mercatoAttivo = Workbook.Mercato;

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            foreach (DataRowView entita in categoriaEntita)
            {
                //si tratta di un'informazione di mercato (tutte le info con _MI e una cifra e visibili)
                DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaInformazione LIKE '%_MI%' AND Visibile = '1' AND IdApplicazione = " + Workbook.IdApplicazione;

                if (informazioni.Count > 0)
                {
                    Range rng = _definedNames.Get(entita["SiglaEntita"], "TEMP_OMI1", Date.SuffissoDATA1).Extend(colOffset: Date.GetOreGiorno(Workbook.DataAttiva));

                    _ws.Range[rng.ToString()].Value = int.Parse(mercatoAttivo.Replace("MI", ""));

                    foreach (DataRowView info in informazioni)
                    {
                        object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                        string informazione = info["SiglaInformazione"].ToString();
                        if (informazione.Equals("PJOLLY_MI") || informazione.Equals("RIFERIMENTO_MERCATO_MI"))
                        {
                            continue;
                        }
                        int row = _definedNames.GetRowByName(siglaEntita, info["SiglaInformazione"]);
                        string mercato = Regex.Match(info["SiglaInformazione"].ToString(), @"_MI\d").Value.Replace("_", "");
                        int col = _definedNames.GetFirstCol() - 2;
                        _ws.Rows[row].EntireRow.Hidden = mercato != mercatoAttivo;

                        //TODO solo per scopi debug: Rimuovere!!!
                        _ws.Rows.Cells[row, col].Value = info["DesInformazione"].ToString() + " " + mercato;
                    }
                }
            }
        }

        protected override void FormattaInformazione(DataRowView info, Excel.Range rngInfo, Excel.Range rngRow, Excel.Range rngData, object testoAlternativo = null)
        {
            base.FormattaInformazione(info, rngInfo, rngRow, rngData, testoAlternativo);
            string siglaInformazione = info["SiglaInformazione"].ToString();
            if (Regex.IsMatch(siglaInformazione, @"OFFERTA_MI\d_G\dE") || Regex.IsMatch(siglaInformazione, @"OFFERTA_MI\d_G\dP"))
            {
                Range rng = new Range(rngData.Address);
                Excel.Validation v = _ws.Range[rng.ToString()].Validation;
                v.Delete();
                v.Add(Type: Excel.XlDVType.xlValidateDecimal,
                    AlertStyle: Excel.XlDVAlertStyle.xlValidAlertStop,
                    Operator: Excel.XlFormatConditionOperator.xlGreaterEqual,
                    Formula1: "0");
                v.IgnoreBlank = false;
                //  v.InputTitle = "Valore";
                // v.InputMessage = "Digitare un valore maggiore o uguale a zero";
                v.ErrorTitle = "Valore non ammesso";
                v.ErrorMessage = "Il valore digitato non è corretto. Sono ammessi solo valori positivi";
                v.ShowError = true;
                v.ShowInput = true;
                Marshal.ReleaseComObject(v);
                v = null;
            }


            //switch (info["SiglaInformazione"].ToString())
            //{ 
            //case "UNIT_COMM":
            //case "RISPETTO_PROG_PREC": break; //info["SiglaInformazione"].Equals("RISPETTO_PROG_PREC")
            //default: //OFFERTA_MI1_G1E //OFFERTA_MI1_G1P 
            //    if (!(info["DesInformazione"].Equals("ACQ/VEN")) && !(info["DesInformazione"].Equals("Codice bilanciamento")) )
            //    {
            //        Range rng = new Range(rngData.Address);
            //        Excel.Validation v = _ws.Range[rng.ToString()].Validation;
            //        v.Delete();
            //        v.Add(Type: Excel.XlDVType.xlValidateDecimal,
            //            AlertStyle: Excel.XlDVAlertStyle.xlValidAlertStop,
            //            Operator: Excel.XlFormatConditionOperator.xlGreaterEqual,
            //            Formula1: "0");
            //        v.IgnoreBlank = false;
            //      //  v.InputTitle = "Valore";
            //       // v.InputMessage = "Digitare un valore maggiore o uguale a zero";
            //        v.ErrorTitle = "Valore non ammesso";
            //        v.ErrorMessage = "Il valore digitato non è corretto. Sono ammessi solo valori positivi";
            //        v.ShowError = true;
            //        v.ShowInput = true;
            //        Marshal.ReleaseComObject(v);
            //        v = null;
            //    }
            //    break;

            //}


            //}       
        }

        

        #endregion
    }
}
