using Iren.PSO.Base;
using System;
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

        //06/02/2017 MOD: nascondo le righe dei mercati non di competenza.
        public override bool Struttura(bool avoidRepositoryUpdate)
        {
            return base.Struttura(avoidRepositoryUpdate);
        }
        /// <summary>
        /// Esegue prima la generazione dei fogli di export, successivamente quella dei fogli di lavoro.
        /// </summary>
        protected override void StrutturaFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.LoadStructure();
                s.HideMarketRows();
            }
        }
        /// <summary>
        /// I label sono diversi quindi viene utilizzato un init label customizzato.
        /// </summary>
        protected override void StrutturaRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.LoadStructure();
        }

        /// <summary>
        /// Aggiorna i dati dei fogli e dei fogli di export.
        /// </summary>
        /// <returns>True se il processo è andato a buon fine.</returns>
        public override bool Dati(bool marketUpdate = true)
        {
            return base.Dati(marketUpdate);
        }
        protected override void DatiFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.UpdateData();
                s.HideMarketRows();
            }
        }

        protected override void DatiRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.UpdateData();
        }

        public override void SetMercatoAttivo()
        {
            Workbook.Mercato = Simboli.GetActiveMarket(DateTime.Now.Hour);
        }

    }
}
