using Iren.PSO.Base;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    public class Aggiorna : Base.Aggiorna
    {
        public Aggiorna()
            : base()
        {

        }

        protected override void StrutturaFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.LoadStructure();
            }

            //AggiornaPrevisioneRiepilogo();
        }
        protected override void StrutturaRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.LoadStructure();
        }

        protected override void DatiFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.UpdateData();
            }

            //AggiornaPrevisioneRiepilogo();
        }
        protected override void DatiRiepilogo()
        {
            Riepilogo main = new Riepilogo();
            main.UpdateData();
        }

        public void AggiornaPrevisioneRiepilogo()
        {
            Riepilogo r = new Riepilogo();
            DataView categoriaEntita = new DataView(Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA]);
            categoriaEntita.RowFilter = "SiglaEntita <> 'UP_TUTTE'";
            foreach (DataRowView entita in categoriaEntita)
            {
                r.AggiornaPrevisione(entita["SiglaEntita"]);
            }
        }
    }

}
