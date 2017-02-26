using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Base
{
    public class Style
    {
        /// <summary>
        /// Imposta tutti i bordi allo stile da applicare al range.
        /// </summary>
        /// <param name="s">Stile.</param>
        /// <param name="colorIndex">Colore del bordo.</param>
        /// <param name="weight">Spessore del bordo.</param>
        public static void SetAllBorders(Excel.Style s, int colorIndex, Excel.XlBorderWeight weight)
        {
            s.Borders.ColorIndex = colorIndex;
            s.Borders.Weight = weight;
            s.Borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            s.Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }
        /// <summary>
        /// Aggiunge al foglio tutti gli stili necessari alla corretta visualizzazione.
        /// </summary>
        public static void StdStyles()
        {
            Excel.Style style;
            try
            {
                style = Workbook.WB.Styles["Top menu GOTO"];
            }
            catch
            {
                style = Workbook.WB.Styles.Add("Top menu GOTO");
                style.Font.Bold = false;
                style.Font.Name = "Verdana";
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 16;

                style = Workbook.WB.Styles.Add("Barra navigazione con date");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 8;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 15;
                style.NumberFormat = "ddd d";
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);

                style = Workbook.WB.Styles.Add("Barra navigazione con nomi");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 8;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 15;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);

                style = Workbook.WB.Styles.Add("Barra titolo entita");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 16;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 15;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlMedium);

                style = Workbook.WB.Styles.Add("Barra della data");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 12;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.NumberFormat = "dddd d mmmm yyyy";
                style.Interior.ColorIndex = 15;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlMedium);

                style = Workbook.WB.Styles.Add("Area grafici");
                style.Font.Size = 10;
                style.Font.Name = "Verdana";
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.NumberFormat = "dddd d mmmm yyyy";
                style.Interior.ColorIndex = 2;
                style.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                style = Workbook.WB.Styles.Add("Area dati");
                style.Font.Size = 10;
                style.Font.Name = "Verdana";
                style.NumberFormat = "#,##0.0;-#,##0.0;-";
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 2;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);

                style = Workbook.WB.Styles.Add("Barra titolo verticale");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 2;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlMedium);

                style = Workbook.WB.Styles.Add("Barra titolo riepilogo");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 37;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlMedium);

                style = Workbook.WB.Styles.Add("Lista entita riepilogo");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 9;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                style = Workbook.WB.Styles.Add("Area dati riepilogo");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 9;
                style.Interior.ColorIndex = 2;
                style.Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);

                style = Workbook.WB.Styles.Add("Lista categorie riepilogo");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 9;
                style.Interior.ColorIndex = 44;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                style = Workbook.WB.Styles.Add("Cella ok");
                style.Font.ColorIndex = 1;
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 9;
                style.Interior.ColorIndex = 4;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                style = Workbook.WB.Styles.Add("Cella non presente");
                style.Font.ColorIndex = 3;
                style.Font.Bold = false;
                style.Font.Name = "Verdana";
                style.Font.Size = 7;
                style.Interior.ColorIndex = 2;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                style = Workbook.WB.Styles.Add("Adjustable");
                style.Font.Color = System.Drawing.Color.Coral;
                //style.Font.Size = 10;
                style.Interior.ColorIndex = 35;
                //style.NumberFormat = "#,##0.0;-#,##0.0;-";
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);
            }
        }
        /// <summary>
        /// Applica tutte le proprietà di stile definite per le applicazioni al range rng.
        /// </summary>
        /// <param name="rng">Range a cui applicare gli stili.</param>
        /// <param name="fontName">Stringa rappresentante il nome del font (i.e. Verdana).</param>
        /// <param name="style">Stringa rappresentante il nome di uno stile predefinito (i.e. Normale).</param>
        /// <param name="merge">Booleano che indica se il range deve essere unito o no.</param>
        /// <param name="bold">Booleano che indica se il font deve essere in grassetto o no.</param>
        /// <param name="fontSize">Intero che indica la dimensione testo.</param>
        /// <param name="align">Oggetto Microsoft.Office.Interop.Excel.XlHAlign per definire il tipo di allineamento orizzontale del range.</param>
        /// <param name="numberFormat">Stringa rappresentante il formato numero della cella.</param>
        /// <param name="foreColor">ColorIndex rappresentante il colore del testo della cella.</param>
        /// <param name="backColor">ColorIndex rappresentante il colore dello sfondo della cella.</param>
        /// <param name="pattern">Oggetto Microsoft.Office.Interop.Excel.XlPattern per definire il tipo di pattern da applicare allo sfondo della cella.</param>
        /// <param name="borders">Stringa che definisce quali bordi e con che spessore debbano essere disegnati. La stringa è nel formato [Top|Left|Bottom|Right|InsideH|InsideV = Thick|Thin|Medium|Hairline, ...].</param>
        /// <param name="orientation">Oggetto Microsoft.Office.Interop.Excel.XlOrientation che definisce che orientazione deve avere il testo del range</param>
        /// <param name="visible">Booleano che indica se il range deve essere visibile oppure no.</param>
        public static void RangeStyle(Excel.Range rng, object fontName = null, object style = null, object merge = null, object bold = null, object fontSize = null, object align = null, object numberFormat = null, object foreColor = null, object backColor = null, object pattern = null, object borders = null, object orientation = null, object visible = null)
        {
            //applica stile per prima cosa
            if (style != null)
                rng.Style = (string)style;

            //Font
            if(fontName != null)
                rng.Font.Name = (string)fontName;
            
            if(bold != null)
                rng.Font.Bold = (bool)bold;
            
            if(fontSize != null)
                rng.Font.Size = (int)fontSize;
            
            if(foreColor != null)
                rng.Font.ColorIndex = (int)foreColor;
            //end Font

            if(merge != null)
                rng.MergeCells = (bool)merge;

            if (align != null)
            {
                rng.HorizontalAlignment = (Excel.XlHAlign)align;
                rng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }

            if(numberFormat != null)
                rng.NumberFormat = (string)numberFormat;

            if (pattern != null)
                rng.Interior.Pattern = (Excel.XlPattern)pattern;

            if(backColor != null)
                rng.Interior.ColorIndex = (int)backColor;

            if(orientation != null)
                rng.Orientation = (Excel.XlOrientation)orientation;

            if(visible != null)
                rng.EntireRow.Hidden = !(bool)visible;

            if (borders != null)
            {
                MatchCollection borderString = Regex.Matches((string)borders, @"(Top|Left|Bottom|Right|InsideH|InsideV)([:=]\w*)?", RegexOptions.IgnoreCase);
                foreach (Match border in borderString)
                {
                    string[] b = Regex.Split(border.Value, @"\s*[:=]\s*");

                    Excel.XlBordersIndex index = Excel.XlBordersIndex.xlEdgeTop;
                    Excel.XlBorderWeight weight = Excel.XlBorderWeight.xlThin;
                    switch (b[0].ToLowerInvariant())
                    {
                        case "top":
                            index = Excel.XlBordersIndex.xlEdgeTop;
                            break;
                        case "left":
                            index = Excel.XlBordersIndex.xlEdgeLeft;
                            break;
                        case "bottom":
                            index = Excel.XlBordersIndex.xlEdgeBottom;
                            break;
                        case "right":
                            index = Excel.XlBordersIndex.xlEdgeRight;
                            break;
                        case "insideh":
                            index = Excel.XlBordersIndex.xlInsideHorizontal;
                            break;
                        case "insidev":
                            index = Excel.XlBordersIndex.xlInsideVertical;
                            break;
                    }
                    if (b.Length == 2)
                    {
                        switch (b[1].ToLowerInvariant())
                        {
                            case "thick":
                                weight = Excel.XlBorderWeight.xlThick;
                                break;
                            case "thin":
                                weight = Excel.XlBorderWeight.xlThin;
                                break;
                            case "medium":
                                weight = Excel.XlBorderWeight.xlMedium;
                                break;
                            case "hairline":
                                weight = Excel.XlBorderWeight.xlHairline;
                                break;
                        }
                    }
                    rng.Borders[index].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rng.Borders[index].Weight = weight;
                }
            }
        }
    }
}
