using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel.Base
{
    public abstract class ACaricaInEmergenza
    {
        public abstract bool RunCarica(object siglaEntita, object siglaAzione, DateTime dataRif);
    }

    /// <summary>
    /// Classe che serve per l'override dell'azione carica dei fogli.
    /// </summary>
    public class CaricaInEmergenza : ACaricaInEmergenza
    {
        /// <summary>
        /// Metodo base di caricamento dei dati.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità per cui caricare i dati.</param>
        /// <param name="siglaAzione">Azione per cui fare il caricamento.</param>
        /// <param name="dataRif">Data su cui fare il caricamento dei dati</param>
        /// <returns></returns>
        public override bool RunCarica(object siglaEntita, object siglaAzione, DateTime dataRif)
        {
            return true;
        }
    }
}
