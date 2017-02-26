using System;
using System.Data;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Base
{
    public class EsportaXML
    {
        public EsportaXML()
        {
        }


        public void RunExport()
        {
            DataTable categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA];
            DataView categorie = Workbook.Repository[DataBase.TAB.CATEGORIA].DefaultView;
            DataView entitaInformazione = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            foreach (DataRow entita in categoriaEntita.Rows)
            {
                object siglaEntita = entita["SiglaEntita"];
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                if (nomeFoglio != "")
                {
                    DefinedNames definedNames = new DefinedNames(nomeFoglio);

                    Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];

                    bool hasData0H24 = definedNames.HasData0H24;

                    entitaInformazione.RowFilter = "(SiglaEntita = '" + siglaEntita + "' OR SiglaEntitaRif = '" + siglaEntita + "') AND Editabile = '1' AND IdApplicazione = " + Workbook.IdApplicazione;

                    DataTable entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA];
                    int intervalloGiorni =
                        (from r in entitaProprieta.AsEnumerable()
                         where r["IdApplicazione"].Equals(Workbook.IdApplicazione) && r["SiglaEntita"].Equals(siglaEntita) && r["SiglaProprieta"].ToString().EndsWith("GIORNI_STRUTTURA")
                         select int.Parse(r["Valore"].ToString())).FirstOrDefault();

                    DateTime dataFine = Workbook.DataAttiva.AddDays(Math.Max(
                        (from r in entitaProprieta.AsEnumerable()
                         where r["IdApplicazione"].Equals(Workbook.IdApplicazione) && r["SiglaEntita"].Equals(siglaEntita) && r["SiglaProprieta"].ToString().EndsWith("GIORNI_STRUTTURA")
                         select int.Parse(r["Valore"].ToString())).FirstOrDefault(), Struct.intervalloGiorni));

                    foreach (DataRowView info in entitaInformazione)
                    {
                        object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                        if (Struct.tipoVisualizzazione == "O")
                        {
                            //prima cella della riga da salvare (non considera Data0H24)
                            Range rng = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], Date.SuffissoDATA1).Extend(colOffset: Date.GetOreIntervallo(dataFine));
                            Handler.StoreEdit(ws.Range[rng.ToString()], 0, true, DataBase.TAB.EXPORT_XML);
                        }
                        else
                        {
                            for (DateTime giorno = Workbook.DataAttiva; giorno <= dataFine; giorno = giorno.AddDays(1))
                            {
                                Range rng = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], Date.GetSuffissoData(giorno)).Extend(colOffset: Date.GetOreGiorno(giorno));
                                Handler.StoreEdit(ws.Range[rng.ToString()], 0, true, DataBase.TAB.EXPORT_XML);
                            }
                        }
                    }
                }
            }

            //preparo l'export


            string directory = Path.Combine(Workbook.Repository.Applicazione["PathDatiComuniEmergenza"].ToString(), Simboli.NomeApplicazione.Replace(" ", ""));
            string fileName = Path.Combine(directory, Simboli.NomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml");

            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            DataTable export = Workbook.Repository[DataBase.TAB.EXPORT_XML];
            export.WriteXml(fileName);

            //svuoto la tabella alla fine dell'utilizzo
            export.Clear();
        }
    }
}
