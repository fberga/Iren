﻿using Iren.PSO.Base;
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
        }
        protected override void DatiRiepilogo()
        {
            Riepilogo main = new Riepilogo();
            main.UpdateData();
        }
    }

}
