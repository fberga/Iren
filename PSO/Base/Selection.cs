using System.Collections.Generic;
using System.Linq;

namespace Iren.PSO.Base
{
    public class Selection
    {
        #region Variabili
        
        private string _rif = "";
        private Dictionary<string, int> _peers = new Dictionary<string, int>();
        
        #endregion

        #region Proprietà

        public string RifAddress { get { return _rif; } }
        public Dictionary<string, int> SelPeers { get { return _peers; } }

        #endregion

        #region Costruttore

        public Selection(string rifAddress, Dictionary<string, int> peers)
        {
            _rif = rifAddress;
            _peers = peers;
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Imposta tutte le selezioni a vuote
        /// </summary>
        /// <param name="ws">Worksheet dove si trova la selezione.</param>
        public void ClearSelections(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            foreach (string cell in SelPeers.Keys)
            {
                double height = ws.Range[cell].RowHeight;
                ws.Range[cell].Value = "\u25CB";
                ws.Range[cell].Font.Size = 15;                
                ws.Range[cell].RowHeight = height;
            }
        }
        /// <summary>
        /// Imposta la selezione in base al valore val.
        /// </summary>
        /// <param name="ws">Worksheet dove si trova la selezione.</param>
        /// <param name="val">Valore da selezionare.</param>
        public void Select(Microsoft.Office.Interop.Excel.Worksheet ws, int val)
        {
            Select(ws, GetByValue(val));
        }
        /// <summary>
        /// Imposta la selezioe in base al range rng selezionato.
        /// </summary>
        /// <param name="ws">Worksheet dove si trova la selezione.</param>
        /// <param name="rng">Range selezionato.</param>
        public void Select(Microsoft.Office.Interop.Excel.Worksheet ws, string rng)
        {
            ws.Range[rng].Value = "\u25CF"; //"\u25C9";
        }
        /// <summary>
        /// Restituisce il range da selezionare in base al valore.
        /// </summary>
        /// <param name="value">Valore.</param>
        /// <returns>Indirizzo in formato A1 del range da selezionare.</returns>
        public string GetByValue(int value)
        {
            return SelPeers.First(kv => kv.Value == value).Key;
        }

        #endregion
    }
}
