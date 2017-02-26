using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Cambio i label e nascondo la riga 6.
    /// </summary>
    class Riepilogo : Base.Riepilogo
    {
        List<DataRowView> _listaEntita = new List<DataRowView>();


        public Riepilogo()
            : base()
        {

        }

        public Riepilogo(Excel.Worksheet ws)
            : base(ws)
        {

        }

        public override void InitLabels()
        {
            base.InitLabels();

            //nascondi quelli non utilizzati
            _ws.Shapes.Item("lbImpianti").Visible = Office.MsoTriState.msoFalse;
            _ws.Shapes.Item("lbElsag").Visible = Office.MsoTriState.msoFalse;

            //sposto i due label sotto
            _ws.Shapes.Item("lbModifica").Top = _ws.Shapes.Item("lbImpianti").Top;
            _ws.Shapes.Item("lbTest").Top = _ws.Shapes.Item("lbElsag").Top;

            //ridimensiono lo sfondo
            _ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoFalse;
            _ws.Shapes.Item("sfondo").Height = (float)(12.5 * _ws.Rows[5].Height);
            _ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoTrue;
        }
        public override void UpdateData()
        {
            _ws.Shapes.Item("lbDataInizio").TextFrame.Characters().Text = Workbook.DataAttiva.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbDataFine").TextFrame.Characters().Text = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("ddd d MMM yyyy");

            //Aggiorno date
            Range rng = new Range(_definedNames.GetRowByName(Date.SuffissoDATA1), _definedNames.GetColFromName("GIORNI"), Struct.intervalloGiorni + 1);

            DateTime giorno = Workbook.DataAttiva;
            foreach (Range row in rng.Rows)
            {
                _ws.Range[row.ToString()].Value = giorno;
                giorno = giorno.AddDays(1);
            }

            //foreach (DataRowView categoria in _categorie)
            //{
            //    _entita.RowFilter = "SiglaEntita <> 'UP_TUTTE' AND SiglaCategoria = '" + categoria["SiglaCategoria"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            //    foreach (DataRowView entita in _entita)
            //        AggiornaPrevisione(entita["SiglaEntita"]);
            //}
        }

        public override void LoadStructure()
        {
            _colonnaInizio = _struttura.colRecap;
            _rigaAttiva = _struttura.rowRecap + 1;

            InitLabels();
            base.Clear();

            _categorie.RowFilter = "Operativa = 1 AND IdApplicazione = " + Workbook.IdApplicazione;
            _entita.RowFilter = "IdApplicazione = " + Workbook.IdApplicazione;

            CreaNomiCelle();
            FormattaRiepilogo();
            FormuleRiepilogo();
            //Se sono in multiscreen lascio il riepilogo alla fine, altrimenti lo riporto all'inizio
            if (Screen.AllScreens.Length == 1)
            {
                _ws.Application.ActiveWindow.SmallScroll(Type.Missing, Type.Missing, _struttura.colRecap - _struttura.colBlock - 1);
            }
        }

        protected override void CreaNomiCelle()
        {
            //inserisco tutte le righe
            _definedNames.AddName(_rigaAttiva++, "TOTALE");
            _definedNames.AddName(_rigaAttiva++, "ENTITA");
            CicloGiorni((oreGiorno, suffissioData, giorno) => 
            {
                _definedNames.AddName(_rigaAttiva++, suffissioData);
            });


            //inserisco tutte le colonne
            _definedNames.AddCol(_colonnaInizio++, "GIORNI");
            foreach (DataRowView categoria in _categorie)
            {
                _entita.RowFilter = "SiglaEntita <> 'UP_TUTTE' AND SiglaCategoria = '" + categoria["SiglaCategoria"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                foreach (DataRowView entita in _entita)
                {
                    _definedNames.AddCol(_colonnaInizio++, entita["SiglaEntita"]);
                    _listaEntita.Add(entita);
                }
            }
            _definedNames.AddCol(_colonnaInizio++, "TOTALE");

            _definedNames.DumpToDataSet();
        }
        protected void FormattaRiepilogo()
        {
            //Titolo in alto
            Range rngTitolo = new Range(_definedNames.GetRowByName("TOTALE"), _definedNames.GetColFromName("GIORNI") + 1, 1, _definedNames.GetColOffsetRiepilogo() - 1);
            Style.RangeStyle(_ws.Range[rngTitolo.ToString()], style: "Barra titolo riepilogo", merge: true, fontSize: 10);
            _ws.Range[rngTitolo.ToString()].Value = "TOTALI";
            _ws.Range[rngTitolo.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            //barra delle date
            Range rngBarraDate = new Range(_definedNames.GetRowByName(Date.SuffissoDATA1), _definedNames.GetColFromName("GIORNI"), Struct.intervalloGiorni + 1);
            Style.RangeStyle(_ws.Range[rngBarraDate.ToString()], style: "Lista entita riepilogo", numberFormat: "dd/MM/yyyy", borders: "[insideh:thin]", bold: false, align: Excel.XlHAlign.xlHAlignCenter);
            _ws.Range[rngBarraDate.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            //compilo i giorni
            int i = 0;
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                _ws.Range[rngBarraDate.Rows[i++].ToString()].Value = giorno;
            });

            //area dati
            Range rngDati = new Range(_definedNames.GetRowByName(Date.SuffissoDATA1), _definedNames.GetColFromName(_listaEntita[0]["SiglaEntita"]), Struct.intervalloGiorni + 1, _definedNames.GetColOffsetRiepilogo() - 1);
            Style.RangeStyle(_ws.Range[rngDati.ToString()], style: "Area dati riepilogo", bold: false, pattern: Excel.XlPattern.xlPatternNone, align: Excel.XlHAlign.xlHAlignCenter, numberFormat: "#,##0;-#,##0;0");                      
            _ws.Range[rngDati.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            bool first = true;
            foreach (DataRowView entita in _listaEntita)
            {
                Range rngEntita = _definedNames.Get("ENTITA", entita["SiglaEntita"]).Extend(Struct.intervalloGiorni + 2);
                Style.RangeStyle(_ws.Range[rngEntita.Rows[0].ToString()], style: "Barra titolo riepilogo", fontSize: 9, borders: "[right:thin" + (!first ? ",left:thin]" : "]"));
                _ws.Range[rngEntita.Rows[0].ToString()].Value = entita["DesEntita"];
                first = false;
            }

            //colonna del totale
            Range rngTotale = _definedNames.Get("ENTITA", "TOTALE").Extend(Struct.intervalloGiorni + 2);
            //titolo
            _ws.Range[rngTotale.Rows[0].ToString()].Value = "TOTALE";
            Style.RangeStyle(_ws.Range[rngTotale.ToString()], style: "Barra titolo riepilogo", fontSize: 9, borders: "[insideh:thin]", bold: true, numberFormat: "#,##0;-#,##0;0");
            //bordi totale
            _ws.Range[rngTotale.Rows[0].ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            _ws.Range[rngTotale.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            Range rngAll = new Range(_definedNames.GetFirstRow(), _definedNames.GetFirstCol(), _definedNames.GetRowOffset(), _definedNames.GetColOffsetRiepilogo());

            _ws.Range[rngAll.ToString()].ColumnWidth = 17;
        }
        protected void FormuleRiepilogo()
        {
            Range rngEntita1 = _definedNames.Get(Date.SuffissoDATA1, _listaEntita[0]["SiglaEntita"]);
            Range rngEntita2 = _definedNames.Get(Date.SuffissoDATA1, _listaEntita.Last()["SiglaEntita"]);

            Range totale = _definedNames.Get(Date.SuffissoDATA1, "TOTALE");
            totale.Extend(Struct.intervalloGiorni + 1);

            _ws.Range[totale.ToString()].Formula = "=SUM(" + rngEntita1.ToString() + ":" + rngEntita2.ToString() + ")";
        }

        public void AggiornaPrevisione(object siglaEntita)
        {
            DateTime dataFine = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni);
            int mainRow = _definedNames.GetRowByName(Date.SuffissoDATA1);
            int mainCol = 0;

            DataView entitaInformazione = new DataView(Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE]);
            entitaInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            Excel.Worksheet ws = Workbook.WB.Sheets[nomeFoglio];
            DefinedNames definedNames = new DefinedNames(nomeFoglio, DefinedNames.InitType.Naming);

            //copio nel main il valore dei totali
            foreach (DataRowView info in entitaInformazione)
            {
                int col = definedNames.GetFirstCol();
                int row = definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], Date.SuffissoDATA1);

                Array rngTotali = ws.Range[Range.GetRange(row, col + 25, Struct.intervalloGiorni)].Value as Array;

                int i = 1;
                for (DateTime giorno = Workbook.DataAttiva; giorno < dataFine; giorno = giorno.AddDays(1))
                {
                    Range[] rngGiornoGas = Sheet.GetRangeGiornoGas(giorno, info, definedNames);

                    Array primoGiorno = ws.Range[rngGiornoGas[0].ToString()].Value as Array;
                    Array secondoGiorno = ws.Range[rngGiornoGas[1].ToString()].Value as Array;

                    if (!(primoGiorno.OfType<double>().Any() || secondoGiorno.OfType<double>().Any()))
                        rngTotali.SetValue(null, i, 1);

                    i++;
                }

                mainCol = _definedNames.GetColFromName(siglaEntita);
                Excel.Range rngMain = Workbook.Main.Range[Range.GetRange(mainRow, mainCol, Struct.intervalloGiorni)];
                rngMain.Value = rngTotali;
            }
        }

        public void SalvaPrevisione()
        {
            DataTable dt = Workbook.Repository[DataBase.TAB.MODIFICA];

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            //DataView entitaInformazione = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita <> 'UP_TUTTE'";
            
            foreach (DataRowView entita in categoriaEntita)
            {
                int col = _definedNames.GetColFromName(entita["SiglaEntita"]);
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    object value = _ws.Range[Range.GetRange(_definedNames.GetRowByName(suffissoData), col)].Value;
                    DataRow newRow = dt.NewRow();

                    newRow["SiglaEntita"] = entita["SiglaEntita"];
                    newRow["SiglaInformazione"] = "CONSUMO_GAS_PREVISIONE_GIORNO";
                    newRow["Data"] = giorno.ToString("yyyyMMdd");
                    newRow["Valore"] = value ?? "";
                    newRow["AnnotaModifica"] = "0";
                    newRow["IdApplicazione"] = Workbook.IdApplicazione;
                    newRow["IdUtente"] = Workbook.IdUtente;

                    dt.Rows.Add(newRow);
                });
            }

            DataBase.SalvaModificheDB();
        }
    }
}
