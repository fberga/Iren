using System;
using System.Data;
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
            : base(ws) {}

        #endregion

        #region Metodi

        /// <summary>
        /// Applica lo stile "Barra titolo verticale" con alcune modifiche e scrive la descrizione dell'entità.
        /// </summary>
        /// <param name="desEntita">Descrizione entità da scrivere.</param>
        protected override void InsertTitoloVerticale(object desEntita)
        {
            DataView informazioni = Base.Workbook.Repository[Base.DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            Base.Range rngTitolo = new Base.Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], Base.Date.GetSuffissoData(_dataInizio)), _struttura.colBlock - _visSelezione - 1, Base.Struct.tipoVisualizzazione == "R" ? _intervalloGiorniMax + 1 : informazioni.Count);

            Excel.Range titoloVert = _ws.Range[rngTitolo.ToString()];
            int infoCount = Base.Struct.tipoVisualizzazione == "R" ? _intervalloGiorniMax + 1 : informazioni.Count;

            Base.Style.RangeStyle(titoloVert, style: "Barra titolo verticale", orientation: infoCount == 1 ? Excel.XlOrientation.xlHorizontal : Excel.XlOrientation.xlVertical, merge: true, fontSize: infoCount == 1 ? 6 : 9, numberFormat: informazioni.Count > 4 ? "ddd d" : "dd");
        }

        #endregion

    }
}
