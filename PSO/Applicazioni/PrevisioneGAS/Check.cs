using Iren.PSO.Base;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Funzione di check.
    /// </summary>
    class Check : Base.Check
    {
        public override CheckOutput ExecuteCheck(Excel.Worksheet ws, DefinedNames definedNames, CheckObj check)
        {
            //Funzione che non centra nulla con i check ma che permette di effettuare il refresh del riepilogo ad ogni azioni che può modificarlo. 
            Aggiorna aggiorna = new Aggiorna();

            aggiorna.AggiornaPrevisioneRiepilogo();

            return new CheckOutput();
        }
    }
}
