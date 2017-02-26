using Iren.PSO.Base;
using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Aggiungo la personalizzazione delle note.
    /// </summary>
    class Sheet : Base.Sheet
    {
        public Sheet(Excel.Worksheet ws)
            : base(ws) { }

        protected override void InsertPersonalizzazioni(object siglaEntita)
        {
            //da classe base il filtro è corretto
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            _ws.Columns[3].Font.Size = 9;

            int col = _definedNames.GetFirstCol();
            siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            int row = _definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio));

            Excel.Range rngPersonalizzazioni = _ws.Range[Range.GetRange(row + 1, col + 25, informazioni.Count - 1)];

            rngPersonalizzazioni.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
            rngPersonalizzazioni.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            rngPersonalizzazioni.Columns[1].ColumnWidth = Struct.cell.width.jolly1;
            rngPersonalizzazioni.WrapText = true;

            //da classe base _dataInizio e _dataFine sono corretti
            for (int i = 1; i < informazioni.Count; i++)
            {
                siglaEntita = informazioni[i]["SiglaEntitaRif"] is DBNull ? informazioni[i]["SiglaEntita"] : informazioni[i]["SiglaEntitaRif"];
                _definedNames.AddName(row + i, siglaEntita, "NOTE", Date.GetSuffissoData(_dataInizio));
                _definedNames.SetEditable(row + i, new Range(row + i, col + 25));
            }
        }
        public override void CaricaInformazioni()
        {
            base.CaricaInformazioni();

            try
            {
                if (DataBase.OpenConnection())
                {
                    string start = DataBase.DataAttiva.ToString("yyyyMMdd");
                    string end = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("yyyyMMdd");

                    DataTable note = DataBase.Select(DataBase.SP.APPLICAZIONE_NOTE, "@SiglaEntita=ALL;@DateFrom="+start+";@DateTo="+end) ?? new DataTable();

                    foreach (DataRow nota in note.Rows)
                    {
                        int row = _definedNames.GetRowByNameSuffissoData(nota["SiglaEntita"], "NOTE", Date.GetSuffissoData(nota["Data"].ToString()));
                        int col = _definedNames.GetFirstCol();
                        _ws.Range[Range.GetRange(row, col + 25)].Value = nota["Note"];
                    }
                } 
            }
            catch (Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "CaricaInformazioni Custom UnitComm: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        public override void UpdateData()
        {
            //cancello tutte le NOTE
            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL ) AND IdApplicazione = " + Workbook.IdApplicazione;

            DateTime dataInizio = DataBase.DataAttiva;
            DateTime dataFine = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni);

            int col = _definedNames.GetFirstCol() + 25;

            foreach (DataRowView entita in categoriaEntita)
            {
                DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];

                CicloGiorni(dataInizio, dataFine, (oreGiorno, suffData, g) =>
                {
                    int row = _definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], suffData);
                    _ws.Range[Range.GetRange(row, col, informazioni.Count)].Value = "";
                });
            }
            base.UpdateData();
        }
    }
}
