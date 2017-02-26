using Iren.PSO.Base;
using System;
using System.Data;
using System.Windows.Forms;
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
            _ws = ws;
            _nomiDefiniti = definedNames;
            _check = check;

            CheckOutput n = new CheckOutput();

            switch (check.Type)
            {
                case 1:
                    n = CheckFunc1();
                    break;
                case 2:
                    n = CheckFunc2();
                    break;
                case 3:
                    n = CheckFunc3();
                    break;
                case 4:
                    n = CheckFunc4();
                    break;
                case 5:
                    n = CheckFunc5();
                    break;
                case 6:
                    n = CheckFunc6();
                    break;
            }

            return n;
        }

        private CheckOutput CheckFunc1()
        {
            Range rngCheck = new Range(_check.Range);

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            TreeNode n = new TreeNode(categoriaEntita[0]["DesEntita"].ToString());
            n.Name = _check.SiglaEntita;

            CheckOutput.CheckStatus status = CheckOutput.CheckStatus.Ok;

            DateTime giorno = Workbook.DataAttiva;
            int oreGiorno = Date.GetOreGiorno(giorno);
            
            string suffissoData = Date.GetSuffissoData(giorno);
            int ora = 1;

            TreeNode nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));

            for (int i = 0; i < rngCheck.ColOffset; i++)
            {
                //caricamento dati                
                object temperaturaMTX = GetObject("CE_MTX", "TEMPERATURA", suffissoData, Date.GetSuffissoOra(ora));
                object temperaturaTTX = GetObject("CE_TTX", "TEMPERATURA", suffissoData, Date.GetSuffissoOra(ora));
                object pressioneMTX = GetObject("CE_MTX", "PRESSIONE", suffissoData, Date.GetSuffissoOra(ora));
                object pressioneTTX = GetObject("CE_TTX", "PRESSIONE", suffissoData, Date.GetSuffissoOra(ora));
                object caricoTermico = GetObject("CT_TORINO", "CARICO_TERMICO_PREVISIONE", suffissoData, Date.GetSuffissoOra(ora));
                object prezzoZonale = GetObject("GRUPPO_TORINO", "PREV_PREZZO", suffissoData, Date.GetSuffissoOra(ora));
                object portataCanale = GetObject("CE_MTX", "PREV_PORTATA", suffissoData, Date.GetSuffissoOra(ora));
                object gruppoFrigo = GetObject("CE_TTX", "GRUPPO_FRIGO", suffissoData, Date.GetSuffissoOra(ora));
                string unitCommMT2R = GetString("UP_MT2R", "UNIT_COMM", suffissoData, Date.GetSuffissoOra(ora));
                string unitCommMT3 = GetString("UP_MT3", "UNIT_COMM", suffissoData, Date.GetSuffissoOra(ora));
                string unitCommTN1 = GetString("UP_TN1", "UNIT_COMM", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispCalorePMaxMT2R = GetDecimal("UP_MT2R", "DISPONIBILITA_CALORE_PMAX", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispCalorePMaxMT3 = GetDecimal("UP_MT3", "DISPONIBILITA_CALORE_PMAX", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispCalorePMaxTN1 = GetDecimal("UP_TN1", "DISPONIBILITA_CALORE_PMAX", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispCalorePMinMT2R = GetDecimal("UP_MT2R", "DISPONIBILITA_CALORE_PMIN", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispCalorePMinMT3 = GetDecimal("UP_MT3", "DISPONIBILITA_CALORE_PMIN", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispCalorePMinTN1 = GetDecimal("UP_TN1", "DISPONIBILITA_CALORE_PMIN", suffissoData, Date.GetSuffissoOra(ora));
                object rampaMT2R = GetObject("UP_MT2R", "RAMPA", suffissoData, Date.GetSuffissoOra(ora));
                object rampaMT3 = GetObject("UP_MT3", "RAMPA", suffissoData, Date.GetSuffissoOra(ora));
                object rampaTN1 = GetObject("UP_TN1", "RAMPA", suffissoData, Date.GetSuffissoOra(ora));
                decimal costoMT2R = GetDecimal("UP_MT2R", "COSTO", suffissoData, Date.GetSuffissoOra(ora));
                decimal costoMT3 = GetDecimal("UP_MT3", "COSTO", suffissoData, Date.GetSuffissoOra(ora));
                decimal costoTN1 = GetDecimal("UP_TN1", "COSTO", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMinMGPMT2R = GetDecimal("UP_MT2R", "PMIN", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMaxMGPMT2R = GetDecimal("UP_MT2R", "PMAX", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMinTernaMT2R = GetDecimal("UP_MT2R", "PMIN_TERNA_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMaxTernaMT2R = GetDecimal("UP_MT2R", "PMAX_TERNA_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMinMGPMT3 = GetDecimal("UP_MT3", "PMIN", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMaxMGPMT3 = GetDecimal("UP_MT3", "PMAX", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMinTernaMT3 = GetDecimal("UP_MT3", "PMIN_TERNA_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMaxTernaMT3 = GetDecimal("UP_MT3", "PMAX_TERNA_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMinMGPTN1 = GetDecimal("UP_TN1", "PMIN", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMaxMGPTN1 = GetDecimal("UP_TN1", "PMAX", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMinTernaTN1 = GetDecimal("UP_TN1", "PMIN_TERNA_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMaxTernaTN1 = GetDecimal("UP_TN1", "PMAX_TERNA_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                //controlli
                if (temperaturaMTX == null)
                {
                    nOra.Nodes.Add("Temperatura Moncalieri assente");
                    errore |= true;
                }
                if (temperaturaMTX != null && (double)temperaturaMTX < -20)
                {
                    nOra.Nodes.Add("Temperatura Moncalieri < soglia minima");
                    errore |= true;
                }
                if (temperaturaMTX != null && (double)temperaturaMTX > 45)
                {
                    nOra.Nodes.Add("Temperatura Moncalieri > soglia massima");
                    errore |= true;
                }
                if (temperaturaTTX == null)
                {
                    nOra.Nodes.Add("Temperatura Torino Nord assente");
                    errore |= true;
                }
                if (temperaturaTTX != null && (double)temperaturaTTX < -20)
                {
                    nOra.Nodes.Add("Temperatura Torino Nord < soglia minima");
                    errore |= true;
                }
                if (temperaturaTTX != null && (double)temperaturaTTX > 45)
                {
                    nOra.Nodes.Add("Temperatura Torino Nord > soglia massima");
                    errore |= true;
                }
                if (pressioneMTX == null)
                {
                    nOra.Nodes.Add("Pressione Moncalieri assente");
                    errore |= true;
                }
                if (pressioneMTX != null && (double)pressioneMTX < 850)
                {
                    nOra.Nodes.Add("Pressione Moncalieri < soglia minima");
                    errore |= true;
                }
                if (pressioneMTX != null && (double)pressioneMTX > 1100)
                {
                    nOra.Nodes.Add("Pressione Moncalieri > soglia massima");
                    errore |= true;
                }
                if (pressioneTTX == null)
                {
                    nOra.Nodes.Add("Pressione Torino Nord assente");
                    errore |= true;
                }
                if (pressioneTTX != null && (double)pressioneTTX < 850)
                {
                    nOra.Nodes.Add("Pressione Torino Nord < soglia minima");
                    errore |= true;
                }
                if (pressioneTTX != null && (double)pressioneTTX > 1100)
                {
                    nOra.Nodes.Add("Pressione Torino Nord > soglia massima");
                    errore |= true;
                }
                if (caricoTermico == null)
                {
                    nOra.Nodes.Add("Carico termico assente");
                    errore |= true;
                }
                if (caricoTermico != null && (double)caricoTermico < 10)
                {
                    nOra.Nodes.Add("Carico termico < soglia minima");
                    errore |= true;
                }
                if (caricoTermico != null && (double)caricoTermico > 2000)
                {
                    nOra.Nodes.Add("Carico termico > soglia massima");
                    errore |= true;
                }
                if (prezzoZonale == null)
                {
                    nOra.Nodes.Add("Prezzo zonale assente");
                    errore |= true;
                }
                if (prezzoZonale != null && (double)prezzoZonale < 0)
                {
                    nOra.Nodes.Add("Prezzo zonale < soglia minima");
                    errore |= true;
                }
                if (prezzoZonale != null && (double)prezzoZonale > 500)
                {
                    nOra.Nodes.Add("Prezzo zonale > soglia massima");
                    errore |= true;
                }
                if (portataCanale == null)
                {
                    nOra.Nodes.Add("Portata canale assente");
                    errore |= true;
                }
                if (portataCanale != null && (
                        ((double)portataCanale < 7 && (unitCommMT2R.Equals("off") || unitCommMT2R.Equals("m") || unitCommMT3.Equals("off") || unitCommMT3.Equals("m")))
                     || ((double)portataCanale < 14 && ((unitCommMT2R.Equals("off") || unitCommMT2R.Equals("m")) && (unitCommMT3.Equals("off") || unitCommMT3.Equals("m"))))
                     || ((double)portataCanale < 36 && unitCommMT2R.Equals("rav") && unitCommMT3.Equals("rav"))
                     || ((double)portataCanale < 25 && ((unitCommMT2R.Equals("rav") && unitCommMT3.Equals("m")) || (unitCommMT3.Equals("m") && unitCommMT3.Equals("rav"))))
                     || ((double)portataCanale < 18 && ((unitCommMT2R.Equals("rav") && (unitCommMT3.Equals("off") || unitCommMT3.Equals("ind"))) || ((unitCommMT2R.Equals("off") || unitCommMT2R.Equals("ind")) && unitCommMT3.Equals("rav"))))))
                {
                    nOra.Nodes.Add("Portata canale < soglia minima");
                    errore |= true;
                }
                if (portataCanale != null && (double)portataCanale > 90)
                {
                    nOra.Nodes.Add("Portata canale > soglia massima");
                    errore |= true;
                }
                if (gruppoFrigo == null)
                {
                    nOra.Nodes.Add("Numero gruppi frigo assente");
                    errore |= true;
                }
                if (gruppoFrigo != null && (double)gruppoFrigo < 0)
                {
                    nOra.Nodes.Add("Numero gruppi frigo < soglia minima");
                    errore |= true;
                }
                if (gruppoFrigo != null && (double)gruppoFrigo > 6)
                {
                    nOra.Nodes.Add("Numero gruppi frigo > soglia massima");
                    errore |= true;
                }
                if (dispCalorePMinMT2R > 0 && (unitCommMT2R == "ind" || unitCommMT2R == "off"))
                {
                    nOra.Nodes.Add("MT2R : disponibilità minima calore > 0");
                    errore |= true;
                }
                if (dispCalorePMinMT3 > 0 && (unitCommMT3 == "ind" || unitCommMT3 == "off"))
                {
                    nOra.Nodes.Add("MT3 : disponibilità minima calore > 0");
                    errore |= true;
                }
                if (dispCalorePMinTN1 > 0 && (unitCommTN1 == "ind" || unitCommTN1 == "off"))
                {
                    nOra.Nodes.Add("TN1 : disponibilità minima calore > 0");
                    errore |= true;
                }
                if (dispCalorePMinMT2R > dispCalorePMaxMT2R)
                {
                    nOra.Nodes.Add("MT2R : disponibilità minima calore > disponibilità massima calore");
                    errore |= true;
                }
                if (dispCalorePMinMT3 > dispCalorePMaxMT3)
                {
                    nOra.Nodes.Add("MT3 : disponibilità minima calore > disponibilità massima calore");
                    errore |= true;
                }
                if (dispCalorePMinTN1 > dispCalorePMaxTN1)
                {
                    nOra.Nodes.Add("TN1 : disponibilità minima calore > disponibilità massima calore");
                    errore |= true;
                }
                if (unitCommMT2R.Equals("rav") && rampaMT2R == null)
                {
                    nOra.Nodes.Add("MT2R : Con assetto rav è necessario inserire il valore di potenza di rampa");
                    errore |= true;
                }
                if (unitCommMT3.Equals("rav") && rampaMT3 == null)
                {
                    nOra.Nodes.Add("MT3 : Con assetto rav è necessario inserire il valore di potenza di rampa");
                    errore |= true;
                }
                if (unitCommTN1.Equals("rav") && rampaTN1 == null)
                {
                    nOra.Nodes.Add("TN1 : Con assetto rav è necessario inserire il valore di potenza di rampa");
                    errore |= true;
                }
                if (costoMT2R == 0)
                {
                    nOra.Nodes.Add("MT2R : Costo marginale assente");
                    attenzione |= true;
                }
                if (costoMT3 == 0)
                {
                    nOra.Nodes.Add("MT3 : Costo marginale assente");
                    attenzione |= true;
                }
                if (costoTN1 == 0)
                {
                    nOra.Nodes.Add("TN1 : Costo marginale assente");
                    attenzione |= true;
                }
                if (pMinTernaMT2R > pMinMGPMT2R && !unitCommMT2R.Equals("rav") && !unitCommMT2R.Equals("off"))
                {
                    nOra.Nodes.Add("MT2R : PMin Terna > PMin MGP");
                    errore |= true;
                }
                if (pMaxMGPMT2R > pMaxTernaMT2R)
                {
                    nOra.Nodes.Add("MT2R : PMax MGP > PMax Terna");
                    errore |= true;
                }
                if (pMinTernaMT3 > pMinMGPMT3 && !unitCommMT3.Equals("rav") && !unitCommMT3.Equals("off"))
                {
                    nOra.Nodes.Add("MT3 : PMin Terna > PMin MGP");
                    errore |= true;
                }
                if (pMaxMGPMT3 > pMaxTernaMT3)
                {
                    nOra.Nodes.Add("MT3 : PMax MGP > PMax Terna");
                    errore |= true;
                }
                if (pMinTernaTN1 > pMinMGPTN1 && !unitCommTN1.Equals("rav") && !unitCommTN1.Equals("off"))
                {
                    nOra.Nodes.Add("TN1 : PMin Terna > Pmin MGP");
                    errore |= true;
                }
                if (pMaxMGPTN1 > pMaxTernaTN1)
                {
                    nOra.Nodes.Add("TN1 : PMax MGP > PMax Terna");
                    errore |= true;
                }
                //fine controlli

                if (errore)
                {
                    ErrorStyle(ref nOra);
                    status = CheckOutput.CheckStatus.Error;
                }
                else if (attenzione)
                {
                    AlertStyle(ref nOra);
                    if (status != CheckOutput.CheckStatus.Error)
                        status = CheckOutput.CheckStatus.Alert;
                }

                nOra.Name = "'" + _ws.Name + "'!" + rngCheck.Columns[i].ToString();

                if (nOra.Nodes.Count > 0)
                    nData.Nodes.Add(nOra);

                string value = errore ? "ERRORE" : attenzione ? "ATTENZ." : "OK";
                _ws.Range[rngCheck.Columns[i].ToString()].Value = value;

                ora = ora < oreGiorno ? ora + 1 : 1;
                if (ora == 1)
                {
                    giorno = giorno.AddDays(1);
                    oreGiorno = Date.GetOreGiorno(giorno);
                    suffissoData = Date.GetSuffissoData(giorno);

                    if (nData.Nodes.Count > 0)
                        n.Nodes.Add(nData);

                    nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));
                }
            }


            //for (int i = 1; i <= rngCheck.ColOffset; i++)
            //{
            //    string suffissoData = Date.GetSuffissoData(DataBase.DataAttiva.AddHours(i - 1));
            //    if (data != DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy"))
            //    {
            //        data = DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy");
            //        if(nData.Nodes.Count > 0)
            //            n.Nodes.Add(nData);

            //        nData = new TreeNode(data);
            //    }

            //    int ora = (i - 1) % Date.GetOreGiorno(suffissoData) + 1;
            //}
            
            if (nData.Nodes.Count > 0)
            {
                n.Nodes.Add(nData);
            }

            if (n.Nodes.Count > 0)
                return new CheckOutput(n, status);

            return new CheckOutput();
        }

        private CheckOutput CheckFunc2()
        {
            //18/01/2016 NEW Check per Tecnoborgo
            //if (_check.SiglaEntita == "UP_INC_TECNOBORGO")
            //    return CheckFunc6();
            //else
            //{
            Range rngCheck = new Range(_check.Range);

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            TreeNode n = new TreeNode(categoriaEntita[0]["DesEntita"].ToString());
            n.Name = _check.SiglaEntita;

            CheckOutput.CheckStatus status = CheckOutput.CheckStatus.Ok;

            DateTime giorno = Workbook.DataAttiva;
            int oreGiorno = Date.GetOreGiorno(giorno);

            string suffissoData = Date.GetSuffissoData(giorno);
            int ora = 1;

            TreeNode nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));

            for (int i = 0; i < rngCheck.ColOffset; i++)
            {
                //caricamento dati
                decimal dispPMax = GetDecimal(_check.SiglaEntita, "DISPONIBILITA_PMAX_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispPMin = GetDecimal(_check.SiglaEntita, "DISPONIBILITA_PMIN_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMax = GetDecimal(_check.SiglaEntita, "PMAX", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMin = GetDecimal(_check.SiglaEntita, "PMIN", suffissoData, Date.GetSuffissoOra(ora));
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                //controlli
                if (pMax > dispPMax)
                {
                    nOra.Nodes.Add("PMax > disponibilità PMax");
                    errore |= true;
                }
                if (pMin < dispPMin)
                {
                    nOra.Nodes.Add("PMin < disponibilità PMin");
                    errore |= true;
                }
                //fine controlli

                if (errore)
                {
                    ErrorStyle(ref nOra);
                    status = CheckOutput.CheckStatus.Error;
                }
                else if (attenzione)
                {
                    AlertStyle(ref nOra);
                    if (status != CheckOutput.CheckStatus.Error)
                        status = CheckOutput.CheckStatus.Alert;
                }

                nOra.Name = "'" + _ws.Name + "'!" + rngCheck.Columns[i].ToString();

                if (nOra.Nodes.Count > 0)
                    nData.Nodes.Add(nOra);

                string value = errore ? "ERRORE" : attenzione ? "ATTENZ." : "OK";
                _ws.Range[rngCheck.Columns[i].ToString()].Value = value;

                ora = ora < oreGiorno ? ora + 1 : 1;
                if (ora == 1)
                {
                    giorno = giorno.AddDays(1);
                    oreGiorno = Date.GetOreGiorno(giorno);
                    suffissoData = Date.GetSuffissoData(giorno);

                    if (nData.Nodes.Count > 0)
                        n.Nodes.Add(nData);

                    nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));
                }
            }

            if (nData.Nodes.Count > 0)
            {
                n.Nodes.Add(nData);
            }

            if (n.Nodes.Count > 0)
                return new CheckOutput(n, status);

            return new CheckOutput();
            //}
        }
        private CheckOutput CheckFunc3()
        {
            Range rngCheck = new Range(_check.Range);

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            TreeNode n = new TreeNode(categoriaEntita[0]["DesEntita"].ToString());
            n.Name = _check.SiglaEntita;

            CheckOutput.CheckStatus status = CheckOutput.CheckStatus.Ok;

            DateTime giorno = Workbook.DataAttiva;
            int oreGiorno = Date.GetOreGiorno(giorno);
            
            string suffissoData = Date.GetSuffissoData(giorno);
            int ora = 1;

            TreeNode nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));

            for (int i = 0; i < rngCheck.ColOffset; i++)
            {
                //caricamento dati
                decimal dispPMaxCC = GetDecimal("GE_GRE1", "DISPONIBILITA_PMAX_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispPMinCC = GetDecimal("GE_GRE1", "DISPONIBILITA_PMIN_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMaxCC = GetDecimal("GE_GRE1", "PMAX", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMinCC = GetDecimal("GE_GRE1", "PMIN", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispPMaxTV = GetDecimal("GE_GRE1", "DISPONIBILITA_PMAX_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispPMinTV = GetDecimal("GE_GRE1", "DISPONIBILITA_PMIN_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMaxTV = GetDecimal("GE_GRE1", "PMAX", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMinTV = GetDecimal("GE_GRE1", "PMIN", suffissoData, Date.GetSuffissoOra(ora));
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                //controlli
                if(pMaxCC + pMaxTV > dispPMaxCC + dispPMaxTV) 
                {
                    nOra.Nodes.Add("PMax > disponibilità PMax");
                    errore |= true;
                }
                if(pMinCC + pMinTV < dispPMinCC + dispPMinTV) 
                {
                    nOra.Nodes.Add("PMin < disponibilità PMin");
                    errore |= true;
                }
                //fine controlli

                if (errore)
                {
                    ErrorStyle(ref nOra);
                    status = CheckOutput.CheckStatus.Error;
                }
                else if (attenzione)
                {
                    AlertStyle(ref nOra);
                    if (status != CheckOutput.CheckStatus.Error)
                        status = CheckOutput.CheckStatus.Alert;
                }

                nOra.Name = "'" + _ws.Name + "'!" + rngCheck.Columns[i].ToString();

                if (nOra.Nodes.Count > 0)
                    nData.Nodes.Add(nOra);

                string value = errore ? "ERRORE" : attenzione ? "ATTENZ." : "OK";
                _ws.Range[rngCheck.Columns[i].ToString()].Value = value;

                ora = ora < oreGiorno ? ora + 1 : 1;
                if (ora == 1)
                {
                    giorno = giorno.AddDays(1);
                    oreGiorno = Date.GetOreGiorno(giorno);
                    suffissoData = Date.GetSuffissoData(giorno);

                    if (nData.Nodes.Count > 0)
                        n.Nodes.Add(nData);

                    nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));
                }
            }

            if (nData.Nodes.Count > 0)
            {
                n.Nodes.Add(nData);
            }

            if (n.Nodes.Count > 0)
                return new CheckOutput(n, status);

            return new CheckOutput();
        }
        private CheckOutput CheckFunc4()
        {
            Range rngCheck = new Range(_check.Range);

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            TreeNode n = new TreeNode(categoriaEntita[0]["DesEntita"].ToString());
            n.Name = _check.SiglaEntita;

            CheckOutput.CheckStatus status = CheckOutput.CheckStatus.Ok;

            DateTime giorno = Workbook.DataAttiva;
            int oreGiorno = Date.GetOreGiorno(giorno);

            string suffissoData = Date.GetSuffissoData(giorno);
            int ora = 1;

            TreeNode nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));

            for (int i = 0; i < rngCheck.ColOffset; i++)
            {
                //caricamento dati                
                object temperatura = GetObject(_check.SiglaEntita, "TEMPERATURA", suffissoData, Date.GetSuffissoOra(ora));
                object pressione = GetObject(_check.SiglaEntita, "PRESSIONE", suffissoData, Date.GetSuffissoOra(ora));
                object umidita = GetObject(_check.SiglaEntita, "UMIDITA", suffissoData, Date.GetSuffissoOra(ora));
                string unitComm = GetString(_check.SiglaEntita, "UNIT_COMM", suffissoData, Date.GetSuffissoOra(ora));
                object rampa = GetObject(_check.SiglaEntita, "RAMPA", suffissoData, Date.GetSuffissoOra(ora));
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                //controlli
                if (temperatura == null)
                {
                    nOra.Nodes.Add("Temperatura assente");
                    errore |= true;
                }
                if (temperatura != null && (double)temperatura < -20)
                {
                    nOra.Nodes.Add("Temperatura < soglia minima");
                    errore |= true;
                }
                if (temperatura != null && (double)temperatura > 45)
                {
                    nOra.Nodes.Add("Temperatura > soglia massima");
                    errore |= true;
                }
                if (pressione == null)
                {
                    nOra.Nodes.Add("Pressione assente");
                    errore |= true;
                }
                if (pressione != null && (double)pressione < 850)
                {
                    nOra.Nodes.Add("Pressione < soglia minima");
                    errore |= true;
                }
                if (pressione != null && (double)pressione > 1100)
                {
                    nOra.Nodes.Add("Pressione > soglia massima");
                    errore |= true;
                }
                if (umidita == null)
                {
                    nOra.Nodes.Add("umidita");
                    errore |= true;
                }
                if (umidita != null && (double)umidita < 5)
                {
                    nOra.Nodes.Add("Umidità < soglia minima");
                    errore |= true;
                }
                if (umidita != null && (double)umidita > 100)
                {
                    nOra.Nodes.Add("Umidità > soglia massima");
                    errore |= true;
                }
                if (unitComm.Equals("rav") && rampa == null)
                {
                    nOra.Nodes.Add("Con assetto rav è necessario inserire il valore di potenza di rampa");
                    errore |= true;
                }
                //fine controlli

                if (errore)
                {
                    ErrorStyle(ref nOra);
                    status = CheckOutput.CheckStatus.Error;
                }
                else if (attenzione)
                {
                    AlertStyle(ref nOra);
                    if (status != CheckOutput.CheckStatus.Error)
                        status = CheckOutput.CheckStatus.Alert;
                }

                nOra.Name = "'" + _ws.Name + "'!" + rngCheck.Columns[i].ToString();

                if (nOra.Nodes.Count > 0)
                    nData.Nodes.Add(nOra);

                string value = errore ? "ERRORE" : attenzione ? "ATTENZ." : "OK";
                _ws.Range[rngCheck.Columns[i].ToString()].Value = value;
                
                ora = ora < oreGiorno ? ora + 1 : 1;
                if (ora == 1)
                {
                    giorno = giorno.AddDays(1);
                    oreGiorno = Date.GetOreGiorno(giorno);
                    suffissoData = Date.GetSuffissoData(giorno);

                    if (nData.Nodes.Count > 0)
                        n.Nodes.Add(nData);

                    nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));
                }
            }

            if (nData.Nodes.Count > 0)
            {
                n.Nodes.Add(nData);
            }

            if (n.Nodes.Count > 0)
                return new CheckOutput(n, status);

            return new CheckOutput();
        }
        private CheckOutput CheckFunc5()
        {
            Range rngCheck = new Range(_check.Range);

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            TreeNode n = new TreeNode(categoriaEntita[0]["DesEntita"].ToString());
            n.Name = _check.SiglaEntita;

            CheckOutput.CheckStatus status = CheckOutput.CheckStatus.Ok;

            DateTime giorno = Workbook.DataAttiva;
            int oreGiorno = Date.GetOreGiorno(giorno);

            string suffissoData = Date.GetSuffissoData(giorno);
            int ora = 1;

            TreeNode nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));

            for (int i = 0; i < rngCheck.ColOffset; i++)
            {
                //caricamento dati                
                object costoOCX = GetObject("UP_OCX", "COSTO", suffissoData, Date.GetSuffissoOra(ora));
                object costoOEX = GetObject("UP_OEX", "COSTO", suffissoData, Date.GetSuffissoOra(ora));
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                //controlli
                if (costoOCX == null || costoOEX == null)
                {
                    nOra.Nodes.Add("Costo marginale assente");
                    attenzione |= true;
                }
                //fine controlli

                if (errore)
                {
                    ErrorStyle(ref nOra);
                    status = CheckOutput.CheckStatus.Error;
                }
                else if (attenzione)
                {
                    AlertStyle(ref nOra);
                    if (status != CheckOutput.CheckStatus.Error)
                        status = CheckOutput.CheckStatus.Alert;
                }

                nOra.Name = "'" + _ws.Name + "'!" + rngCheck.Columns[i].ToString();

                if (nOra.Nodes.Count > 0)
                    nData.Nodes.Add(nOra);

                string value = errore ? "ERRORE" : attenzione ? "ATTENZ." : "OK";
                _ws.Range[rngCheck.Columns[i].ToString()].Value = value;

                ora = ora < oreGiorno ? ora + 1 : 1;
                if (ora == 1)
                {
                    giorno = giorno.AddDays(1);
                    oreGiorno = Date.GetOreGiorno(giorno);
                    suffissoData = Date.GetSuffissoData(giorno);

                    if (nData.Nodes.Count > 0)
                        n.Nodes.Add(nData);

                    nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));
                }
            }

            if (nData.Nodes.Count > 0)
            {
                n.Nodes.Add(nData);
            }

            if (n.Nodes.Count > 0)
                return new CheckOutput(n, status);

            return new CheckOutput();
        }
        //18/01/2017 NEW check per Tecnogorgo
        private CheckOutput CheckFunc6()
        {
            Range rngCheck = new Range(_check.Range);

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            TreeNode n = new TreeNode(categoriaEntita[0]["DesEntita"].ToString());
            n.Name = _check.SiglaEntita;

            CheckOutput.CheckStatus status = CheckOutput.CheckStatus.Ok;

            DateTime giorno = Workbook.DataAttiva;
            int oreGiorno = Date.GetOreGiorno(giorno);

            string suffissoData = Date.GetSuffissoData(giorno);
            int ora = 1;

            TreeNode nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));

            for (int i = 0; i < rngCheck.ColOffset; i++)
            {
                //caricamento dati    
                decimal dispPMax = GetDecimal(_check.SiglaEntita, "DISPONIBILITA_PMAX_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispPMin = GetDecimal(_check.SiglaEntita, "DISPONIBILITA_PMIN_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal consultivoUltimo = GetDecimal(_check.SiglaEntita, "TEMP_PROG1", suffissoData, Date.GetSuffissoOra(ora));
                decimal dispPMAXUltimo = GetDecimal(_check.SiglaEntita, "TEMP_PROG2", suffissoData, Date.GetSuffissoOra(ora));
                decimal ausiliari = GetDecimal(_check.SiglaEntita, "AUSILIARI_ASSETTO1", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMax = GetDecimal(_check.SiglaEntita, "PMAX", suffissoData, Date.GetSuffissoOra(ora));
                decimal pMin = GetDecimal(_check.SiglaEntita, "PMIN", suffissoData, Date.GetSuffissoOra(ora));
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                //controlli
                if (Math.Max(1, (dispPMAXUltimo - ausiliari) * 0.25m) < Math.Abs(Math.Max(0, (dispPMAXUltimo - ausiliari) -
                    //04/02/2017 ENH: modifica specifiche
                    pMax)))
                {
                    nOra.Nodes.Add("Programma molto diverso dalla disponibilità");
                    //04/02/2017 ENH: modifica specifiche
                    attenzione |= true;
                }
                if (dispPMax != dispPMAXUltimo)
                {
                    nOra.Nodes.Add("Variazione di disponibilità");
                    attenzione |= true;
                }
                if (pMax > dispPMax)
                {
                    nOra.Nodes.Add("PMax > disponibilità PMax");
                    errore |= true;
                }
                if (pMin < dispPMin)
                {
                    nOra.Nodes.Add("PMin < disponibilità PMin");
                    errore |= true;
                }
                //fine controlli

                if (errore)
                {
                    ErrorStyle(ref nOra);
                    status = CheckOutput.CheckStatus.Error;
                }
                else if (attenzione)
                {
                    AlertStyle(ref nOra);
                    if (status != CheckOutput.CheckStatus.Error)
                        status = CheckOutput.CheckStatus.Alert;
                }

                nOra.Name = "'" + _ws.Name + "'!" + rngCheck.Columns[i].ToString();

                if (nOra.Nodes.Count > 0)
                    nData.Nodes.Add(nOra);

                string value = errore ? "ERRORE" : attenzione ? "ATTENZ." : "OK";
                _ws.Range[rngCheck.Columns[i].ToString()].Value = value;

                ora = ora < oreGiorno ? ora + 1 : 1;
                if (ora == 1)
                {
                    giorno = giorno.AddDays(1);
                    oreGiorno = Date.GetOreGiorno(giorno);
                    suffissoData = Date.GetSuffissoData(giorno);

                    if (nData.Nodes.Count > 0)
                        n.Nodes.Add(nData);

                    nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));
                }
            }

            if (nData.Nodes.Count > 0)
            {
                n.Nodes.Add(nData);
            }

            if (n.Nodes.Count > 0)
                return new CheckOutput(n, status);

            return new CheckOutput();
        }
    }
}
