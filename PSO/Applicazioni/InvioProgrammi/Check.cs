using Iren.PSO.Base;
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

            System.DateTime giorno = Workbook.DataAttiva;
            int oreGiorno = Date.GetOreGiorno(giorno);

            string suffissoData = Date.GetSuffissoData(giorno);
            int ora = 1;

            TreeNode nData = new TreeNode(giorno.ToString("dd-MM-yyyy"));

            for (int i = 0; i < rngCheck.ColOffset; i++)
            {
                //caricamento dati
                decimal programmaQ1_MT2R = GetDecimal("UP_MT2R", "PROGRAMMAQ1_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal programmaQ2_MT2R = GetDecimal("UP_MT2R", "PROGRAMMAQ2_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal programmaQ3_MT2R = GetDecimal("UP_MT2R", "PROGRAMMAQ3_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal programmaQ4_MT2R = GetDecimal("UP_MT2R", "PROGRAMMAQ4_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal psMinAccettata_MT2R = GetDecimal("UP_MT2R", "PSMIN_ACCETTATA", suffissoData, Date.GetSuffissoOra(ora));

                decimal programmaQ1_MT3 = GetDecimal("UP_MT3", "PROGRAMMAQ1_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal programmaQ2_MT3 = GetDecimal("UP_MT3", "PROGRAMMAQ2_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal programmaQ3_MT3 = GetDecimal("UP_MT3", "PROGRAMMAQ3_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal programmaQ4_MT3 = GetDecimal("UP_MT3", "PROGRAMMAQ4_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal psMinAccettata_MT3 = GetDecimal("UP_MT3", "PSMIN_ACCETTATA", suffissoData, Date.GetSuffissoOra(ora));
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                //controlli
                if (programmaQ1_MT2R > 0 && programmaQ1_MT2R < psMinAccettata_MT2R)
                {
                    nOra.Nodes.Add("Programma < PMin");
                    attenzione |= true;
                }
                if (programmaQ2_MT2R > 0 && programmaQ2_MT2R < psMinAccettata_MT2R)
                {
                    nOra.Nodes.Add("Programma < PMin");
                    attenzione |= true;
                }
                if (programmaQ3_MT2R > 0 && programmaQ3_MT2R < psMinAccettata_MT2R)
                {
                    nOra.Nodes.Add("Programma < PMin");
                    attenzione |= true;
                }
                if (programmaQ4_MT2R > 0 && programmaQ4_MT2R < psMinAccettata_MT2R)
                {
                    nOra.Nodes.Add("Programma < PMin");
                    attenzione |= true;
                }
                if (programmaQ1_MT3 > 0 && programmaQ1_MT3 < psMinAccettata_MT3)
                {
                    nOra.Nodes.Add("Programma < PMin");
                    attenzione |= true;
                }
                if (programmaQ2_MT3 > 0 && programmaQ2_MT3 < psMinAccettata_MT3)
                {
                    nOra.Nodes.Add("Programma < PMin");
                    attenzione |= true;
                }
                if (programmaQ3_MT3 > 0 && programmaQ3_MT3 < psMinAccettata_MT3)
                {
                    nOra.Nodes.Add("Programma < PMin");
                    attenzione |= true;
                }
                if (programmaQ4_MT3 > 0 && programmaQ4_MT3 < psMinAccettata_MT3)
                {
                    nOra.Nodes.Add("Programma < PMin");
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
        private CheckOutput CheckFunc2()
        {
            Range rngCheck = new Range(_check.Range);

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            TreeNode n = new TreeNode(categoriaEntita[0]["DesEntita"].ToString());
            n.Name = _check.SiglaEntita;
            TreeNode nData = new TreeNode();
            string data = "";

            CheckOutput.CheckStatus status = CheckOutput.CheckStatus.Ok;

            for (int i = 1; i <= rngCheck.ColOffset; i++)
            {
                string suffissoData = Date.GetSuffissoData(DataBase.DataAttiva.AddHours(i - 1));
                if (data != DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy"))
                {
                    data = DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy");
                    if(nData.Nodes.Count > 0)
                        n.Nodes.Add(nData);

                    nData = new TreeNode(data);
                }

                int ora = (i - 1) % Date.GetOreGiorno(suffissoData) + 1;

                //caricamento dati                
                decimal programmaQ1 = GetDecimal(_check.SiglaEntita, "PROGRAMMAQ1_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal programmaQ2 = GetDecimal(_check.SiglaEntita, "PROGRAMMAQ2_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal programmaQ3 = GetDecimal(_check.SiglaEntita, "PROGRAMMAQ3_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal programmaQ4 = GetDecimal(_check.SiglaEntita, "PROGRAMMAQ4_" + Workbook.Mercato, suffissoData, Date.GetSuffissoOra(ora));
                decimal psMinAccettata = GetDecimal(_check.SiglaEntita, "PSMIN_ACCETTATA", suffissoData, Date.GetSuffissoOra(ora));
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                //controlli
                if (programmaQ1 > 0 && programmaQ1 < psMinAccettata)
                {
                    nOra.Nodes.Add("Programma < PMin");
                    attenzione |= true;
                }
                if (programmaQ2 > 0 && programmaQ2 < psMinAccettata)
                {
                    nOra.Nodes.Add("Programma < PMin");
                    attenzione |= true;
                }
                if (programmaQ3 > 0 && programmaQ3 < psMinAccettata)
                {
                    nOra.Nodes.Add("Programma < PMin");
                    attenzione |= true;
                }
                if (programmaQ4 > 0 && programmaQ4 < psMinAccettata)
                {
                    nOra.Nodes.Add("Programma < PMin");
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

                nOra.Name = "'" + _ws.Name + "'!" + rngCheck.Columns[i - 1].ToString();

                if (nOra.Nodes.Count > 0)
                    nData.Nodes.Add(nOra);

                string value = errore ? "ERRORE" : attenzione ? "ATTENZ." : "OK";
                _ws.Range[rngCheck.Columns[i - 1].ToString()].Value = value;
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
