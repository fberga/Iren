using Iren.PSO.Base;
using System;
using System.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    class Sheet : Base.Sheet
    {
        public Sheet(Excel.Worksheet ws) : base(ws)
        {

        }

        /// <summary>
        /// Carica il profilo PQNR in seguito al caricamento delle informazioni.
        /// </summary>
        public override void CaricaInformazioni()
        {
            base.CaricaInformazioni();
            
            //profili PQNR
            if (_ws.Name == "Iren Termo")
            {
                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                DataView entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].DefaultView;
                DataView entitaRampa = Workbook.Repository[DataBase.TAB.ENTITA_RAMPA].DefaultView;
                categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND IdApplicazione = " + Workbook.IdApplicazione;

                foreach (DataRowView entita in categoriaEntita)
                {
                    DateTime dataFine;
                    entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_STRUTTURA' AND IdApplicazione = " + Workbook.IdApplicazione;
                    
                    if (entitaProprieta.Count > 0)
                        dataFine = _dataInizio.AddDays(double.Parse("" + entitaProprieta[0]["Valore"]));
                    else
                        dataFine = _dataInizio.AddDays(Struct.intervalloGiorni);

                    double pRif =
                        (from r in Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].AsEnumerable()
                         where r["IdApplicazione"].Equals(Workbook.IdApplicazione) && r["SiglaEntita"].Equals(entita["SiglaEntita"])
                            && r["SiglaProprieta"].Equals("SISTEMA_COMANDI_PRIF")
                         select Double.Parse(r["Valore"].ToString())).FirstOrDefault();

                    int oreIntervallo = Date.GetOreIntervallo(dataFine);

                    Range rngPQNR = _definedNames.Get(entita["SiglaEntita"], "PQNR_PROFILO", Date.SuffissoDATA1).Extend(colOffset: oreIntervallo);

                    if (_ws.Range[rngPQNR.Columns[0].ToString()].Value != null)
                    {
                        int assetti = Workbook.Repository[DataBase.TAB.ENTITA_ASSETTO].AsEnumerable().Count(r => r["SiglaEntita"].Equals(entita["SiglaEntita"]));

                        double[] pMin = new double[oreIntervallo];
                        for (int i = 0; i < pMin.Length; i++) pMin[i] = double.MaxValue;

                        for (int i = 0; i < assetti; i++)
                        {
                            Range rngPmin = _definedNames.Get(entita["SiglaEntita"], "PMIN_TERNA_ASSETTO" + (i + 1), Date.SuffissoDATA1).Extend(colOffset: oreIntervallo);
                            for (int j = 0; j < oreIntervallo; j++)
                                pMin[j] = Math.Min(pMin[j], (double)(_ws.Range[rngPmin.Columns[j].ToString()].Value ?? 0d));
                        }

                        object[,] valori = new object[24, oreIntervallo];
                        for (int i = 0; i < oreIntervallo; i++)
                        {
                            pMin[i] = pMin[i] < pRif ? pRif : pMin[i];
                            entitaRampa.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaRampa = '" + _ws.Range[rngPQNR.Columns[i].ToString()].Value + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                            if (entitaRampa.Count > 0)
                            {
                                for (int j = 0; j < 24; j++)
                                {
                                    if (entitaRampa[0]["Q" + (j + 1)] != DBNull.Value)
                                    {
                                        valori[j, i] = Math.Round(((int)entitaRampa[0]["Q" + (j + 1)]) * pRif / pMin[i]);
                                    }
                                }
                            }
                        }
                        Range rngPQNRVal = _definedNames.Get(entita["SiglaEntita"], "PQNR1", Date.SuffissoDATA1).Extend(rowOffset: 24, colOffset: oreIntervallo);
                        _ws.Range[rngPQNRVal.ToString()].Value = valori;
                    }
                }
            }
        }

        protected override void InsertGrafici()
        {
            base.InsertGrafici();

            foreach (Excel.ChartObject chart in _ws.ChartObjects())
            {
                chart.Chart.PlotVisibleOnly = true;
            }
        }
    }
}
