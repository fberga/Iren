using Iren.PSO.Base;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.PSO.Applicazioni
{
    class Riepilogo : Base.Riepilogo
    {
        public Riepilogo()
            : base()
        {

        }

        public Riepilogo(Excel.Worksheet ws)
            : base(ws)
        {

        }

        /// <summary>
        /// Inizializza i label con parametri customizzati.
        /// </summary>
        public override void InitLabels()
        {
            base.InitLabels();

            _ws.Shapes.Item("lbMercato").Visible = Office.MsoTriState.msoTrue;
            _ws.Shapes.Item("lbMercato").Top = _ws.Shapes.Item("lbDataInizio").Top + _ws.Shapes.Item("lbDataInizio").Height + (float)(_ws.Rows[5].Height / 2);

            Handler.ChangeMercatoAttivo(Workbook.Mercato);
            //sposto i due label sotto

            _ws.Shapes.Item("lbUtente").Top = _ws.Shapes.Item("lbMercato").Top + _ws.Shapes.Item("lbMercato").Height + (float)_ws.Rows[5].Height;
            _ws.Shapes.Item("lbSQLServer").Top = _ws.Shapes.Item("lbUtente").Top + (float)(_ws.Rows[5].Height * 2);
            _ws.Shapes.Item("lbImpianti").Top = _ws.Shapes.Item("lbUtente").Top + (float)(_ws.Rows[5].Height * 4);
            _ws.Shapes.Item("lbElsag").Top = _ws.Shapes.Item("lbUtente").Top + (float)(_ws.Rows[5].Height * 6);
            _ws.Shapes.Item("lbModifica").Top = _ws.Shapes.Item("lbUtente").Top + (float)(_ws.Rows[5].Height * 8);
            _ws.Shapes.Item("lbTest").Top = _ws.Shapes.Item("lbUtente").Top + (float)(_ws.Rows[5].Height * 10);

            //ridimensiono lo sfondo
            _ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoFalse;
            _ws.Shapes.Item("sfondo").Height = (float)(19.5 * _ws.Rows[5].Height);
            _ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoTrue;
        }
    }
}
