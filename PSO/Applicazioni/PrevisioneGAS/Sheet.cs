using Iren.PSO.Base;
using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Aggiungo la personalizzazione delle note.
    /// </summary>
    class Sheet : Base.Sheet
    {
        public Sheet(Excel.Worksheet ws)
            : base(ws) 
        { 
        
        }

        protected override void InsertPersonalizzazioni(object siglaEntita)
        {
            //da classe base il filtro è corretto
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            _ws.Columns[3].Font.Size = 9;

            int col = _definedNames.GetFirstCol();
            int row = _definedNames.GetRowByName(siglaEntita, "T");

            //metto cella con scritta totale            
            //Excel.Range title = _ws.Range[Range.GetRange(row, col + 25)];
            //title.Value = "TOTALE";

            Excel.Range rngPersonalizzazioni = _ws.Range[Range.GetRange(row + 2, col + 25, _intervalloGiorniMax)];

            rngPersonalizzazioni.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

            //title.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            //Style.RangeStyle(title, bold: true, backColor: 8, align: Excel.XlHAlign.xlHAlignCenter);
            
            rngPersonalizzazioni.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            rngPersonalizzazioni.Columns[1].ColumnWidth = Struct.cell.width.jolly1;
            Style.RangeStyle(rngPersonalizzazioni, fontName: "Verdana", fontSize: 9, bold: true);
            
            int i = 1;
            //filtro giusto da classe base
            DataView entitaInformazione = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
            foreach (DataRowView info in entitaInformazione)
            {
                CicloGiorni(_dataInizio, _dataInizio.AddDays(Struct.intervalloGiorni - 1), (oreGiorno, suffissoData, giorno) =>
                {
                    //row = _definedNames.GetRowByName(siglaEntita,  "T");

                    //int gasDayStart = TimeZone.CurrentTimeZone.IsDaylightSavingTime(giorno) ? 7 : 6;
                    //int remainingHours = 24 - Date.GetOreGiorno(giorno) + gasDayStart;

                    //Range rng1 = new Range(row + i + 1, _definedNames.GetColData1H1() + gasDayStart - 1, 1, 25 - gasDayStart + 1);
                    //Range rng2 = new Range(row + i + 2, _definedNames.GetColData1H1(), 1, remainingHours - 1);

                    Range[] giornoGas = GetRangeGiornoGas(giorno, info, _definedNames);

                    rngPersonalizzazioni.Cells[i, 1].Formula = "=SUM(" + giornoGas[0].ToString() + ") + SUM(" + giornoGas[1].ToString() + ")";
                    i++;
                });   
            }
        }

        public static Range[] GetRangeGiornoGas(DateTime giorno, DataRowView info, DefinedNames definedNames)
        {
            int row = definedNames.GetRowByName(info["SiglaEntita"], info["SiglaInformazione"], Date.GetSuffissoData(giorno));

            int gasDayStart = TimeZone.CurrentTimeZone.IsDaylightSavingTime(giorno) ? 7 : 6;
            int remainingHours = 24 - Date.GetOreGiorno(giorno) + gasDayStart;

            Range rng1 = new Range(row, definedNames.GetColData1H1() + gasDayStart, 1, Date.GetOreGiorno(giorno) - gasDayStart);
            Range rng2 = new Range(row + 1, definedNames.GetColData1H1(), 1, remainingHours);

            return new Range[] { rng1, rng2 };
        }

        //protected override void InsertGrafici()
        //{
        //    DataView grafici = Workbook.Repository[DataBase.TAB.ENTITA_GRAFICO].DefaultView;
        //    DataView graficiInfo = Workbook.Repository[DataBase.TAB.ENTITA_GRAFICO_INFORMAZIONE].DefaultView;

        //    int i = 1;
        //    int col = _definedNames.GetColData1H1();
        //    int colOffset = _definedNames.GetColOffset(_dataFine) - (_struttura.visData0H24 ? 1 : 0);
        //    foreach (DataRowView grafico in grafici)
        //    {
        //        SplashScreen.UpdateStatus("Genero grafici");
        //        string name = DefinedNames.GetName(grafico["SiglaEntita"], "GRAFICO" + i++, Struct.tipoVisualizzazione == "V" ? Date.GetSuffissoData(_dataInizio) : "");

        //        Range rngGrafico = new Range(_definedNames.GetRowByName(name), col, 1, colOffset);
        //        Excel.Range xlRngGrafico = _ws.Range[rngGrafico.ToString()];
        //        xlRngGrafico.Merge();
        //        xlRngGrafico.Style = "Area grafici";
        //        xlRngGrafico.RowHeight = 200;
        //        Excel.Chart chart = _ws.ChartObjects().Add(Left: xlRngGrafico.Left, Top: xlRngGrafico.Top + 1, Width: xlRngGrafico.Width, Height: xlRngGrafico.Height - 2).Chart;

        //        chart.Parent.Name = name;

        //        chart.Axes(Excel.XlAxisType.xlCategory).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
        //        chart.Axes(Excel.XlAxisType.xlValue).HasMajorGridlines = false;
        //        chart.Axes(Excel.XlAxisType.xlValue).HasMinorGridlines = false;
        //        chart.Axes(Excel.XlAxisType.xlValue).MinorTickMark = Excel.XlTickMark.xlTickMarkOutside;
        //        chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Name = "Verdana";
        //        chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Size = 11;
        //        chart.Axes(Excel.XlAxisType.xlValue).TickLabels.NumberFormat = "general";

        //        chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
        //        chart.HasDataTable = false;
        //        chart.DisplayBlanksAs = Excel.XlDisplayBlanksAs.xlNotPlotted;
        //        chart.ChartGroups(1).GapWidth = 0;
        //        chart.ChartGroups(1).Overlap = 100;
        //        chart.ChartArea.Border.ColorIndex = 1;
        //        chart.ChartArea.Border.Weight = 3;
        //        chart.ChartArea.Border.LineStyle = 0;
        //        chart.PlotVisibleOnly = false;

        //        chart.PlotArea.Top = chart.ChartArea.Height;

        //        chart.PlotArea.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

        //        string rowFilter = graficiInfo.RowFilter;
        //        graficiInfo.RowFilter = rowFilter + " AND SiglaGrafico = '" + grafico["SiglaGrafico"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;

        //        foreach (DataRowView info in graficiInfo)
        //        {
        //            CicloGiorni(_dataInizio, _dataInizio.AddDays(Struct.intervalloGiorni), (oreGiorno, suffissoData, giorno) =>
        //            {
        //                Range[] rngGiorno = GetRangeGiornoGas(giorno, info, _definedNames);

        //                Array a1 = _ws.Range[rngGiorno[0].ToString()].Value as Array;
        //                Array a2 = _ws.Range[rngGiorno[1].ToString()].Value as Array;

        //                Array tot = Array.CreateInstance(typeof(object), 1, Date.GetOreGiorno(giorno));
        //                Array.Copy(a1, tot, a1.Length);
        //                Array.Copy(a2, 1, tot, a1.Length, a2.Length);
                        

        //                Range rngDati = new Range(_definedNames.GetRowByNameSuffissoData(grafico["SiglaEntita"], info["SiglaInformazione"], Date.GetSuffissoData(giorno)), col, 1, Date.GetOreGiorno(giorno));
        //                Excel.Series serie = chart.SeriesCollection().NewSeries();
        //                serie.Name = giorno.ToString("dd/MM/yyyy");
        //                serie.Values = tot;
        //                serie.ChartType = (Excel.XlChartType)info["ChartType"];
        //                //serie.Interior.ColorIndex = info["InteriorColor"];
        //                //serie.Border.ColorIndex = info["BorderColor"];
        //                //serie.Border.Weight = info["BorderWeight"];
        //                //serie.Border.LineStyle = info["BorderLineStyle"];
        //                //serie.AxisGroup = (Excel.XlAxisGroup)info["AxisGroup"];
        //            });
        //        }

        //        graficiInfo.RowFilter = rowFilter;
        //    }
        //}

        protected override void InsertGrafici()
        {
            base.InsertGrafici();

            Excel.ChartObjects charts = _ws.ChartObjects();
            foreach (Excel.ChartObject chart in charts)
            {
                chart.Chart.Axes(Excel.XlAxisType.xlValue).TickLabels.NumberFormat = "[>=1000]#,##0.0,\"K\";0.0";
            }
        }

        public override void AggiornaGrafici()
        {
            if (_ws.ChartObjects().Count > 0)
            {
                ((Excel._Worksheet)_ws).Calculate();
                Excel.ChartObjects charts = _ws.ChartObjects();
                foreach (Excel.ChartObject chart in charts)
                {
                    int col;
                    if (chart.Name.Contains("DATA"))
                    {
                        col = _definedNames.GetColFromDate(chart.Name.Split(Simboli.UNION[0]).Last());
                    }
                    else
                    {
                        col = _definedNames.GetColFromDate();
                    }
                    int row = _definedNames.GetRowByName(chart.Name);
                    Excel.Range rng = _ws.Range[Range.GetRange(row, col)];
                    int i = 0;
                    foreach (Excel.Series s in chart.Chart.SeriesCollection())
                    {
                        s.Name = Workbook.DataAttiva.AddDays(i++).ToString("dd/MM/yyyy");
                    }
                    AggiornaGrafici(chart.Chart, rng.MergeArea);
                    //chart.Chart.Refresh();
                }
            }
        }
        /// <summary>
        /// Allinea il grafico al range in modo da far combaciare la barra delle ordinate con la prima colonna dell'area dati. Per far questo calcola la dimensione in punti dei label di ordinata e sposta di conseguenza l'area del grafico.
        /// </summary>
        /// <param name="chart">Microsoft.Office.Interop.Excel.Chart da aggiornare.</param>
        /// <param name="rigaGrafico">Microsoft.Office.Interop.Excel.Range a cui il grafico appartiene.</param>
        private void AggiornaGrafici(Excel.Chart chart, Excel.Range rigaGrafico)
        {
            SplashScreen.UpdateStatus("Aggiorno grafici " + chart.Name);

            chart.Refresh();
            //resize dell'area del grafico per adattarla alle ore
            using (Graphics grfx = Graphics.FromImage(new Bitmap(1, 1)))
            {
                grfx.PageUnit = GraphicsUnit.Point;
                grfx.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                float sizeMax = float.MinValue;
                SizeF tmpSize;
                double val = chart.Axes(Excel.XlAxisType.xlValue).MinimumScale;

                //controllo anche il fondo scala: se cambia l'ordine di grandezza excel lascia lo spazio nel label come se ci fosse!!
                while (val <= chart.Axes(Excel.XlAxisType.xlValue).MaximumScale)
                {
                    string tmpval = "" 
                        + (val >= 1000 ? Math.Round(val / 1000.0, 1) : val) 
                        + ((val / 1000.0) % 1 == 0 ? ",0" : "") 
                        + (val >= 1000 ? "K" : "");

                    tmpval = tmpval.Replace(".", ",");

                    tmpSize = grfx.MeasureString(tmpval, new Font(chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Name, (float)chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Size));
                    sizeMax = Math.Max(sizeMax, tmpSize.Width);

                    val += chart.Axes(Excel.XlAxisType.xlValue).MajorUnit;
                }

                //MANTENERE ORDINE DI QUESTE ISTRUZIONI
                chart.ChartArea.Left = rigaGrafico.Left - Math.Ceiling(sizeMax) - 7;        //sposto a destra il grafico
                chart.ChartArea.Width = rigaGrafico.Width + Math.Ceiling(sizeMax) + 4;      //aumento la larghezza del grafico
                Excel.PlotArea plotArea = chart.PlotArea;
                try
                {
                    plotArea.InsideLeft = 0d;                                               //allineo il grafico al bordo sinistro dell'area esterna al grafico
                }
                catch { }
                plotArea.Width = chart.ChartArea.Width + 3;                                 //aumento la larghezza dell'area esterna al grafico
                Marshal.ReleaseComObject(plotArea);
                plotArea = null;

                bool start = TimeZone.CurrentTimeZone.IsDaylightSavingTime(Workbook.DataAttiva);
                bool end = TimeZone.CurrentTimeZone.IsDaylightSavingTime(Workbook.DataAttiva.AddDays(Struct.intervalloGiorni));

                if (!start || end)
                    chart.ChartArea.Width -= _ws.Range[Range.GetRange(1, _definedNames.GetColFromDate(Date.SuffissoDATA1, Date.GetSuffissoOra(25)))].Width;
            }
            chart.Refresh();
        }
    }
}
