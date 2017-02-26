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
    }
}
