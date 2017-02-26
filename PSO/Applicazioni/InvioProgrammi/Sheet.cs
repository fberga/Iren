using Iren.PSO.Base;
using System;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Al carica informazioni viene aggiunta la funzione per aggiornare i colori di sfondo delle celle per evidenziare le variazioni dai mercati precedenti. Inoltre, elimina il titolo verticale.
    /// </summary>
    public class Sheet : Base.Sheet
    {
        #region Variabili

        DefinedNames _definedNamesSheetMercato = new DefinedNames(Workbook.Mercato);
        //string _mercatoPrec = Simboli.GetMercatoPrec();
        Excel.Worksheet _wsMercato;

        #endregion

        #region Costruttori

        public Sheet(Excel.Worksheet ws) 
            : base(ws) 
        {
            _wsMercato = Workbook.Sheets[Workbook.Mercato];
        
        }
        
        #endregion

        #region Metodi

        protected override void InsertTitoloVerticale(object desEntita)
        {
            base.InsertTitoloVerticale(desEntita);

            //rimuovo la scritta
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            Range rngTitolo = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _struttura.colBlock - _visSelezione - 1, informazioni.Count);

            Excel.Range titoloVert = _ws.Range[rngTitolo.ToString()];
            titoloVert.Value = null;
        }

        public void AggiornaColori()
        {
            if (Workbook.Mercato != "MSD1")
            {
                SplashScreen.UpdateStatus("Aggiorno colori");

                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

                foreach (DataRowView entita in categoriaEntita)
                {
                    informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaTipologiaInformazione <> 'CHECK' AND Visibile = '1' AND IdApplicazione = " + Workbook.IdApplicazione;
                    foreach (DataRowView info in informazioni)
                    {
                        object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                        Range rng = _definedNames.Get(siglaEntita, info["SiglaInformazione"], Date.SuffissoDATA1).Extend(colOffset: Date.GetOreGiorno(Workbook.DataAttiva));
                        string quarter = Regex.Match(info["SiglaInformazione"].ToString(), @"Q\d").Value;
                        quarter = quarter == "" ? "Q1" : quarter;

                        var rif =
                            (from r in categoriaEntita.Table.AsEnumerable()
                             where r["IdApplicazione"].Equals(Workbook.IdApplicazione) && r["SiglaEntita"].Equals(siglaEntita)
                             select new { SiglaEntita = r["Gerarchia"] is DBNull ? r["SiglaEntita"] : r["Gerarchia"], Riferimento = r["Riferimento"] }).First();

                        Range rngMercato = new Range(_definedNamesSheetMercato.GetRowByName(rif.SiglaEntita, "UM", "T") + 2, _definedNamesSheetMercato.GetColFromName("RIF" + rif.Riferimento, "PROGRAMMA" + quarter)).Extend(rowOffset: Date.GetOreGiorno(Workbook.DataAttiva));

                        for (int j = 0; j < rngMercato.Rows.Count; j++)
                            _ws.Range[rng.Columns[j].ToString()].Interior.ColorIndex = _wsMercato.Range[rngMercato.Rows[j].ToString()].DisplayFormat.Interior.ColorIndex;
                    }
                }
            }
        }

        public override void CaricaInformazioni()
        {
            base.CaricaInformazioni();
            AggiornaColori();
        }
        
        #endregion
    }
}
