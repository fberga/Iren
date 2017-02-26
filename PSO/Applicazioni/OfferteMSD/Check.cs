using Iren.PSO.Base;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Funzioni di check personalizzate.
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
                decimal ePPA = GetDecimal(_check.SiglaEntita, "PEM", suffissoData, Date.GetSuffissoOra(ora));
                decimal ePSMaxAccettata = GetDecimal(_check.SiglaEntita, "PSMAX_ACCETTATA", suffissoData, Date.GetSuffissoOra(ora));
                decimal ePSMinAccettata = GetDecimal(_check.SiglaEntita, "PSMIN_ACCETTATA", suffissoData, Date.GetSuffissoOra(ora));

                decimal offG0V = GetDecimal(_check.SiglaEntita, "OFFERTA_MSD_G0VE", suffissoData, Date.GetSuffissoOra(ora));
                decimal offG0A = GetDecimal(_check.SiglaEntita, "OFFERTA_MSD_G0AE", suffissoData, Date.GetSuffissoOra(ora));
                decimal offG1V = GetDecimal(_check.SiglaEntita, "OFFERTA_MSD_G1VE", suffissoData, Date.GetSuffissoOra(ora));
                decimal offG1A = GetDecimal(_check.SiglaEntita, "OFFERTA_MSD_G1AE", suffissoData, Date.GetSuffissoOra(ora));
                decimal offG2V = GetDecimal(_check.SiglaEntita, "OFFERTA_MSD_G2VE", suffissoData, Date.GetSuffissoOra(ora));
                decimal offG2A = GetDecimal(_check.SiglaEntita, "OFFERTA_MSD_G2AE", suffissoData, Date.GetSuffissoOra(ora));
                decimal offG3V = GetDecimal(_check.SiglaEntita, "OFFERTA_MSD_G3VE", suffissoData, Date.GetSuffissoOra(ora));
                decimal offG3A = GetDecimal(_check.SiglaEntita, "OFFERTA_MSD_G3AE", suffissoData, Date.GetSuffissoOra(ora));
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                //controlli
                if (ePPA > ePSMaxAccettata)
                {
                    nOra.Nodes.Add("Programma di produzione superiore alla PMax Terna");
                    errore |= true;
                }
                if (ePSMaxAccettata < ePPA && ePPA < ePSMinAccettata && ePPA > 0)
                {
                    nOra.Nodes.Add("Programma di produzione non coerente con PMin-PMax Terna");
                    errore |= true;
                }
                if (offG0V < 0 || offG0A < 0 || offG1V < 0 || offG1A < 0 || offG2V < 0 || offG2A < 0 || offG3V < 0 || offG3A < 0)
                {
                    nOra.Nodes.Add("Offerta < 0");
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


        protected override decimal GetDecimal(object siglaEntita, object siglaInformazione, object suffissoData, object suffissoOra)
        {
            Range rng = _nomiDefiniti.Get(siglaEntita, siglaInformazione, suffissoData, suffissoOra);

            if (_ws.Range[rng.ToString()].EntireRow.Hidden)
                return 0;
            else
                return base.GetDecimal(rng);
        }

    }
}
