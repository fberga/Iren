using Iren.PSO.Base;
using System;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Aggiunge un azione custom alla modifica manuale. Questa modifica scrive i dati nel foglio nascosto che corrisponde al mercato attivo.
    /// </summary>
    public class Modifica : Base.Modifica
    {
        public override void Range(object Sh, Excel.Range Target)
        {
            if (Workbook.CategorySheets.Contains(Target.Worksheet))
            {
                //Se la funzione scrive in altre celle, ricordarsi di disabilitare gli handler per la modifica delle celle
                //Workbook.WB.SheetChange -= Handler.StoreEdit;
                Workbook.RemoveStdStoreEdit();
                Workbook.WB.SheetChange -= this.Range;


                Excel.Worksheet ws = Target.Worksheet;
                Excel.Worksheet wsMercato = Workbook.Sheets[Workbook.Mercato];

                bool wasProtected = wsMercato.ProtectContents;
                if (wasProtected)
                {
                    wsMercato.Unprotect(Workbook.Password);
                    ws.Unprotect(Workbook.Password);
                }

                DefinedNames definedNames = new DefinedNames(ws.Name);
                DefinedNames definedNamesMercato = new DefinedNames(Workbook.Mercato);

                DataTable entita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA];

                string[] ranges = Target.Address.Split(',');

                foreach (string range in ranges)
                {
                    Range rng = new Range(range);

                    foreach (Range cell in rng.Cells)
                    {
                        string[] parts = definedNames.GetNameByAddress(cell.StartRow, cell.StartColumn).Split(Simboli.UNION[0]);

                        string siglaEntita = parts[0];
                        string siglaInformazione = parts[1];
                        string suffissoData = parts[2];
                        string suffissoOra = parts[3];

                        var rif =
                        (from r in entita.AsEnumerable()
                         where r["IdApplicazione"].Equals(Workbook.IdApplicazione) && r["SiglaEntita"].Equals(siglaEntita)
                         select new { SiglaEntita = r["Gerarchia"] is DBNull ? r["SiglaEntita"] : r["Gerarchia"], Riferimento = r["Riferimento"] }).First();

                        string quarter = Regex.Match(siglaInformazione, @"Q\d").Value;
                        quarter = quarter == "" ? "Q1" : quarter;

                        Range rngMercato = new Range(definedNamesMercato.GetRowByName(rif.SiglaEntita, "UM", "T") + 2, definedNamesMercato.GetColFromName("RIF" + rif.Riferimento, "PROGRAMMA" + quarter));
                        rngMercato.StartRow += (Date.GetOraFromSuffissoOra(suffissoOra) - 1);

                        wsMercato.Range[rngMercato.ToString()].Value = ws.Range[cell.ToString()].Value;
                        ws.Range[cell.ToString()].Interior.ColorIndex = wsMercato.Range[rngMercato.ToString()].DisplayFormat.Interior.ColorIndex;
                    }
                }

                if (wasProtected)
                {
                    wsMercato.Protect(Workbook.Password);
                    ws.Protect(Workbook.Password);
                }

                //Se la funzione scrive in altre celle, ricordarsi di riabilitare gli handler per la modifica delle celle
                //Workbook.WB.SheetChange += Handler.StoreEdit;
                Workbook.AddStdStoreEdit();
                Workbook.WB.SheetChange += this.Range;
            }
        }
    }
}
