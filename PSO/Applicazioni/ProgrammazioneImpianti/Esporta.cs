using Iren.PSO.Base;
using System;
using System.Data;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;


namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Funzioni di esportazione custom.
    /// </summary>
    public class Esporta : AEsporta
    {
        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif, string[] mercati)
        {
            DataView entitaAzione = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            if (entitaAzione.Count == 0)
                return false;

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            object codiceRUP = categoriaEntita[0]["CodiceRUP"];

            DataView entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].DefaultView;
            entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'IMP_COD_IF' AND IdApplicazione = " + Workbook.IdApplicazione;
            object codiceIF = entitaProprieta[0]["Valore"];

            DataView entitaAzioneInformazione = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
            entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            DefinedNames definedNames = new DefinedNames(nomeFoglio);

            switch (siglaAzione.ToString())
            {
                case "E_MP_MGP":
                    DataTable dt = new DataTable("E_MP_MGP")
                    {
                        Columns =
                        {
                            {"Campo1", typeof(string)},
                            {"Campo2", typeof(string)},
                            {"UP", typeof(string)},
                            {"Campo3", typeof(string)},
                            {"Data", typeof(string)},
                            {"Ora", typeof(string)},
                            {"Informazione", typeof(string)},
                            {"Valore", typeof(string)}
                        }
                    };

                    string suffissoData = Date.GetSuffissoData(dataRif);
                    foreach (DataRowView info in entitaAzioneInformazione)
                    {
                        object siglaEntitaRif = (info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"]);
                        
                        Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];
                        Range range = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], suffissoData);
                        range.Extend(0, definedNames.GetDayOffset(suffissoData) - 1);
                        Excel.Range rng = ws.Range[range.ToString()];

                        object[,] tmpVal = rng.Value;
                        object[] values = tmpVal.Cast<object>().ToArray();

                        for (int i = 0, length = values.Length; i < length; i++)
                        {
                            DataRow row = dt.NewRow();

                            row["Campo1"] = nomeFoglio == "Iren Termo" ? "AHRP" : "AIHRP";
                            row["Campo2"] = "Prod";
                            row["UP"] = codiceIF;
                            if (definedNames.IsDefined(siglaEntitaRif, "UNIT_COMM"))
                                row["Campo3"] = "17";
                            else
                                row["Campo3"] = "NA";
                            row["Data"] = dataRif.ToString("yyyy/MM/dd");
                            row["Ora"] = i + 1;
                            row["Informazione"] = info["SiglaInformazione"].Equals("PMAX") ? "Pmax" : "Pmin";
                            row["Valore"] = values[i] ?? 0;

                            dt.Rows.Add(row);
                        }
                    }

                    string pathStr = PreparePath(Workbook.GetUsrConfigElement("pathExportMP_MGP"));

                    if (Directory.Exists(pathStr))
                    {
                        if (!ExportToCSV(System.IO.Path.Combine(pathStr, "AEM_" + (nomeFoglio == "Iren Termo" ? "AHRP_" : "AIHRP_") + codiceIF + "_" + dataRif.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfffffff") + ".csv"), dt))
                            return false;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Il percorso '" + pathStr + "' non è raggiungibile.", Simboli.NomeApplicazione, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return false;
                    }

                    break;
            }
            return true;
        }
    }
}
