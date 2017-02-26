using Iren.PSO.Base;
using System;
using System.Data;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
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
            entitaAzioneInformazione.RowFilter = "SiglaEntitaRif = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            DefinedNames definedNames = new DefinedNames(nomeFoglio, DefinedNames.InitType.Naming);

            Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];

            switch (siglaAzione.ToString())
            {
                case "E_UNIT_COMM":
                    DataTable dt = new DataTable("E_UNIT_COMM")
                    {
                        Columns =
                        {
                            {"Campo1", typeof(string)},
                            {"Campo2", typeof(string)},
                            {"UP", typeof(string)},
                            {"Campo3", typeof(string)},
                            {"Data", typeof(string)},
                            {"Ora", typeof(string)},
                            {"Campo4", typeof(string)},
                            {"UnitComm", typeof(string)},
                            {"Campo5", typeof(string)}
                        }
                    };

                    string suffissoData = Date.GetSuffissoData(dataRif);
                    int oreData = Date.GetOreGiorno(dataRif);
                    foreach (DataRowView info in entitaAzioneInformazione)
                    {
                        object siglaEntitaRif = (info["SiglaEntitaRif"] is DBNull ? siglaEntita: info["SiglaEntitaRif"]);

                        Range rng = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], suffissoData).Extend(colOffset: oreData);

                        //object[,] values = ws.Range[rng.ToString()].Value;
                        //bool empty = true;
                        //foreach (object value in values)
                        //{
                        //    if(value != null) 
                        //    {
                        //        empty = false;
                        //        break;
                        //    }
                        //}

                        //if (!empty)
                        //{
                            for (int i = 0; i < rng.Columns.Count; i++)
                            {
                                DataRow row = dt.NewRow();

                                row["Campo1"] = "ASSET";
                                row["Campo2"] = "Produzione";
                                row["UP"] = codiceIF;
                                row["Campo3"] = "NA";
                                row["Data"] = dataRif.ToString("dd/MM/yyyy");
                                row["Ora"] = (i + 1).ToString("00") + ".00";
                                row["Campo4"] = "ASSETTO";
                                row["UnitComm"] = ws.Range[rng.Columns[i].ToString()].Value;
                                row["Campo5"] = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                                dt.Rows.Add(row);
                            }
                        //}
                    }

                    string pathStr = PreparePath(Workbook.GetUsrConfigElement("pathCaricatoreImpianti"));

                    if (Directory.Exists(pathStr))
                    {
                        if (dt.AsEnumerable().Any(r => r["UnitComm"] != DBNull.Value)
                            && ExportToCSV(System.IO.Path.Combine(pathStr, "AEM_ASSET_" + codiceIF + "_" + dataRif.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfffffff") + ".csv"), dt))
                            return true;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Il percorso '" + pathStr + "' non è raggiungibile.", Simboli.NomeApplicazione, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    }
                    break;
            }
            return false;
        }
    }
}
