using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Iren.PSO.Forms
{
    public class OfferteMIHelper
    {

        public static Dictionary<string,int> GetGOTODictionary(string siglaEntita, string siglaInformazione, DefinedNames dn) 
        {
            Dictionary<string, int> gotoDictionary = new Dictionary<string, int>();

            DataView definizioneOfferta = Workbook.Repository[DataBase.TAB.DEFINIZIONE_OFFERTA].DefaultView;

            definizioneOfferta.RowFilter = "SiglaEntita ='" + siglaEntita + "' AND SiglaInformazione = '" + siglaInformazione + "' AND IdMercato = " + Workbook.Mercato.Substring(2, Workbook.Mercato.Length - 2);

            if (definizioneOfferta.Count == 0)
                return null;

            DataTable entitaInformazione = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE];

            foreach (DataRowView offerta in definizioneOfferta)
            {
                string desInformazioneCombo = entitaInformazione.AsEnumerable()
                    .Where(r => r["SiglaEntita"].Equals(offerta["SiglaEntita"])
                             && (r["SiglaEntitaRif"] is DBNull || r["SiglaEntitaRif"].Equals(offerta["SiglaEntitaCombo"]))
                             && r["SiglaInformazione"].Equals(offerta["SiglaInformazioneCombo"]))
                    .Select(r => r["DesInformazione"].ToString())
                    .FirstOrDefault();

                object entitaCalcolo = offerta["SiglaEntitaCalcolo"] is DBNull ? offerta["SiglaEntitaCombo"] : offerta["SiglaEntitaCalcolo"];
                object infoCalcolo = offerta["SiglaInformazioneCalcolo"] is DBNull ? offerta["SiglaInformazioneCombo"] : offerta["SiglaInformazioneCalcolo"];

                gotoDictionary.Add(desInformazioneCombo, dn.GetRowByName(entitaCalcolo, infoCalcolo));
            }

            return gotoDictionary;
        }
    }
}
