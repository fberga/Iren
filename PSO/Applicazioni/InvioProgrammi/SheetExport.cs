using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Crea i fogli di export.
    /// </summary>
    class SheetExport : Base.ASheet
    {
        #region Variabili

        private static Dictionary<string, DateTime> _dataCaricaStruttura = new Dictionary<string, DateTime>();

        protected Excel.Worksheet _ws;
        protected DefinedNames _definedNames;
        protected DefinedNames _definedNamesMercatoPrec;
        protected int _rigaAttiva;
        protected string _mercato;
        protected int _appID;

        #endregion

        #region Costruttori

        public SheetExport(Excel.Worksheet ws)
        {
            _ws = ws;
            _mercato = ws.Name;
            _appID = Workbook.Repository[DataBase.TAB.MERCATI].AsEnumerable()
                .Where(r => r["DesMercato"].Equals(_mercato))
                .Select(r => (int)r["IdApplicazioneMercato"])
                .FirstOrDefault();

            AggiornaParametriSheet();

            _definedNames = new DefinedNames(_mercato);
            if (_mercato != "MSD1")
                _definedNamesMercatoPrec = new DefinedNames(Simboli.GetMercatoPrec(_mercato));
        }

        #endregion

        #region Proprietà

        public DateTime DataCaricamentoStruttura
        {
            get
            {
                if (_dataCaricaStruttura.ContainsKey(_mercato))
                    return _dataCaricaStruttura[_mercato];
                else
                    return DateTime.MinValue;
            }
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Aggiorna i parametri applicazione contenuti nella tabella APPLICAZIONE.
        /// </summary>
        protected void AggiornaParametriSheet()
        {
            _struttura = new Struct();

            _struttura.rigaBlock = (int)Workbook.Repository.Applicazione["RowBlocco"];
            _struttura.rigaGoto = (int)Workbook.Repository.Applicazione["RowGoto"];
            _struttura.colBlock = 2;

        }
        /// <summary>
        /// Cancella il contenuto del foglio.
        /// </summary>
        private void Clear()
        {            
            if (_ws.ChartObjects().Count > 0)
                _ws.ChartObjects().Delete();

            _ws.Rows.ClearContents();
            _ws.Rows.ClearComments();
            _ws.Rows.FormatConditions.Delete();
            _ws.Rows.EntireRow.Hidden = false;
            _ws.Rows.UnMerge();
            _ws.Rows.Style = "Normal";

            _ws.Rows.RowHeight = Struct.cell.height.normal;
            _ws.Columns.ColumnWidth = Struct.cell.width.dato;

            _ws.Rows["1:" + (_struttura.rigaBlock - 1)].RowHeight = Struct.cell.height.empty;

            _ws.Rows[_struttura.rigaGoto].RowHeight = Struct.cell.height.normal;

            _ws.Columns[1].ColumnWidth = Struct.cell.width.empty;

            if(_ws.Visible == Excel.XlSheetVisibility.xlSheetVisible)
            {
                ((Excel._Worksheet)_ws).Activate();
                _ws.Application.ActiveWindow.FreezePanes = false;
                _ws.Cells[_struttura.rigaBlock, 1].Select();
                _ws.Application.ActiveWindow.FreezePanes = true;
                Workbook.Main.Select();
            }
            _ws.Application.ScreenUpdating = false;
        }
        /// <summary>
        /// Inizializza la barra del titolo.
        /// </summary>
        protected void InitBarraNavigazione()
        {
            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;

            Excel.Range gotoBar = _ws.Range[_ws.Cells[2, 2], _ws.Cells[_struttura.rigaGoto + 1, categoriaEntita.Count + 3]];
            gotoBar.Style = "Top menu GOTO";
            gotoBar.BorderAround2(Weight: Excel.XlBorderWeight.xlMedium, Color: 1);

            int i = 3;
            foreach (DataRowView entita in categoriaEntita)
            {
                Excel.Range rng = _ws.Cells[_struttura.rigaGoto, i++];
                rng.Value = entita["DesEntitaBreve"];
                rng.Style = "Barra navigazione con nomi";
            }
        }
        /// <summary>
        /// Crea tutte le colonne.
        /// </summary>
        private void InitColumns()
        {
            //definisco tutte le colonne
            DataTable categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA];

            //Calcolo il massimo numero di entità da mettere affiancate
            int maxElementCount =
                (from r in categoriaEntita.AsEnumerable()
                 where r["IdApplicazione"].Equals(_appID) && r["Gerarchia"] != DBNull.Value
                 group r by r["Gerarchia"] into g
                 select g.Count()).Max();

            int colonnaAttiva = _struttura.colBlock;
            for (int i = 0; i < maxElementCount; i++)
            {
                colonnaAttiva++;
                for (int j = 0; j < 4; j++)
                    _definedNames.AddCol(colonnaAttiva++, "RIF" + (i + 1), "PROGRAMMAQ" + (j + 1));
            }
        }
        /// <summary>
        /// Carica la struttura.
        /// </summary>
        public override void LoadStructure()
        { 
            SplashScreen.UpdateStatus("Creo struttura " + _mercato);

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "Gerarchia = '' OR Gerarchia IS NULL AND IdApplicazione = " + _appID;

            //if (DataCaricamentoStruttura != Workbook.DataAttiva)
            //{
                Clear();
                InitBarraNavigazione();
            //}

            InitColumns();

            _rigaAttiva = _struttura.rigaBlock + 1;

            foreach (DataRowView entita in categoriaEntita)
                InitBloccoEntita(entita);

            _definedNames.DumpToDataSet();

            //if (DataCaricamentoStruttura != Workbook.DataAttiva)
                CaricaInformazioni();

            if (_dataCaricaStruttura.ContainsKey(_mercato))
                _dataCaricaStruttura[_mercato] = Workbook.DataAttiva;
            else
                _dataCaricaStruttura.Add(_mercato, Workbook.DataAttiva);
        }
        /// <summary>
        /// Inizializza il blocco entità.
        /// </summary>
        /// <param name="entita">Riga con i dati dell'entità.</param>
        protected void InitBloccoEntita(DataRowView entita)
        {
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND IdApplicazione = " + _appID;
            CreaNomiCelle(entita["SiglaEntita"]);
            
            //if (DataCaricamentoStruttura != Workbook.DataAttiva)
            //{
                FormattaBloccoEntita(entita["SiglaEntita"], entita["DesEntita"], entita["CodiceRUP"]);
            //}

        }
        /// <summary>
        /// Crea i nomi delle celle.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entitaà.</param>
        protected void CreaNomiCelle(object siglaEntita)
        {
            _definedNames.AddName(_rigaAttiva, siglaEntita, "T");
            _rigaAttiva += 2;
            _definedNames.AddName(_rigaAttiva, siglaEntita, "DATA");
            _rigaAttiva += 2;
            _definedNames.AddName(_rigaAttiva, siglaEntita, "UM", "T");
            _rigaAttiva += Date.GetOreGiorno(Workbook.DataAttiva) + 5;
        }
        /// <summary>
        /// Formatta il blocco entità.
        /// </summary>
        /// <param name="siglaEntita">Sigla entità.</param>
        /// <param name="desEntita">Descrizione.</param>
        /// <param name="codiceRUP">Codice RUP.</param>
        protected void FormattaBloccoEntita(object siglaEntita, object desEntita, object codiceRUP)
        {
            Range rngMercatoPrec = new Range();
            //Titolo
            Range rng = new Range(_definedNames.GetRowByName(siglaEntita, "T"), _struttura.colBlock, 1, 10);
            Style.RangeStyle(_ws.Range[rng.ToString()], fontSize: 12, merge: true, bold: true, align: Excel.XlHAlign.xlHAlignCenter, borders: "[top:medium,right:medium,bottom:medium,left:medium]");
            _ws.Range[rng.ToString()].Value = "PROGRAMMA A 15 MINUTI " + desEntita;
            _ws.Range[rng.ToString()].RowHeight = 25;

            //Data
            rng = new Range(_definedNames.GetRowByName(siglaEntita, "DATA"), _struttura.colBlock, 1, 5);
            Style.RangeStyle(_ws.Range[rng.ToString()], fontSize: 10, bold: true, align: Excel.XlHAlign.xlHAlignCenter, borders: "[top:medium,right:medium,bottom:medium,left:medium,insidev:medium]", numberFormat: "dd/MM/yyyy");
            _ws.Range[rng.ToString()].RowHeight = 18;
            _ws.Range[rng.Columns[0].ToString()].Value = "Data";
            _ws.Range[rng.Columns[1, 3].ToString()].Merge();
            _ws.Range[rng.Columns[1].ToString()].Value = Workbook.DataAttiva;
            _ws.Range[rng.Columns[4].ToString()].Value = _mercato;

            //Tabella
            DataTable categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA];
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            List<DataRow> entitaRif =
                (from r in categoriaEntita.AsEnumerable()
                 where r["IdApplicazione"].Equals(Workbook.IdApplicazione) && r["Gerarchia"].Equals(siglaEntita)
                 select r).ToList();
            
            bool hasEntitaRif = entitaRif.Count > 0;
            int numEntita = Math.Max(entitaRif.Count, 1);

            rng = new Range(_definedNames.GetRowByName(siglaEntita, "UM", "T"), _struttura.colBlock, 1, 5 * numEntita);
            for (int i = 0; i < numEntita; i++)
            {
                informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND Visibile = '1' " + (hasEntitaRif ? "AND SiglaEntitaRif = '" + entitaRif[i]["SiglaEntita"] + "'" : "") + " AND IdApplicazione = " + _appID;
                
                //range grande come tutta la tabella
                rng = new Range(_definedNames.GetRowByName(siglaEntita, "UM", "T"), _definedNames.GetColFromName("RIF" + (i + 1), "PROGRAMMAQ1") - 1, Date.GetOreGiorno(Workbook.DataAttiva) + 2, 5);

                Style.RangeStyle(_ws.Range[rng.ToString()], borders: "[top:medium,right:medium,bottom:medium,left:medium,insideH:thin,insideV:thin]", align: Excel.XlHAlign.xlHAlignCenter, numberFormat: "general");
                Style.RangeStyle(_ws.Range[rng.Rows[1, rng.Rows.Count - 1].Columns[0].ToString()], backColor: 15, bold: true, align: Excel.XlHAlign.xlHAlignLeft);
                Style.RangeStyle(_ws.Range[rng.Rows[0].ToString()], backColor: 15, bold: true, fontSize: 11);
                Style.RangeStyle(_ws.Range[rng.Rows[1].ToString()], backColor: 15, bold: true);
                _ws.Range[rng.Rows[0].Columns[1, rng.Columns.Count - 1].ToString()].Merge();
                if (hasEntitaRif)
                    _ws.Range[rng.Rows[0].ToString()].Value = new object[] { "UM", entitaRif[i]["CodiceRUP"] is DBNull ? entitaRif[i]["DesEntita"] : entitaRif[i]["CodiceRUP"] };
                else
                    _ws.Range[rng.Rows[0].ToString()].Value = new object[] { "UM", codiceRUP is DBNull ? desEntita : codiceRUP };

                for (int h = 1; h <= Date.GetOreGiorno(Workbook.DataAttiva); h++)
                    _ws.Range[rng.Columns[0].Rows[h + 1].ToString()].Value = "Ora " + h;

                var isOrario = informazioni
                    .OfType<DataRowView>()
                    .Any(r => r["SiglaInformazione"].ToString().StartsWith("PROGRAMMA_"));

                if (!isOrario)
                {
                    for (int j = 0; j < 4; j++)
                        _ws.Range[rng.Rows[1].Columns[j + 1].ToString()].Value = 15 * j + "-" + 15 * (j+1);
                }
                else
                    _ws.Range[rng.Cells[1,1].ToString()].Value = "0-60";

                //TODO controllare che non ci siano problemi
                if (_mercato != "MSD1")
                {
                    string mercatoPrec = Simboli.GetMercatoPrec(_mercato);
                    //calcolo il range nel foglio del mercato precedente (non è detto che siano nella stessa posizione (anche se non ha senso che non lo siano...))
                    rngMercatoPrec = new Range(_definedNamesMercatoPrec.GetRowByName(siglaEntita, "UM", "T"), _definedNamesMercatoPrec.GetColFromName("RIF" + (i + 1), "PROGRAMMAQ1") - 1, Date.GetOreGiorno(Workbook.DataAttiva) + 2, 5);

                    Excel.FormatCondition condGreater = _ws.Range[rng.Rows[2, rng.Rows.Count - 1].Columns[1, 4].ToString()].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: "=" + rng.Cells[2, 1] + " > '" + mercatoPrec + "'!" + rngMercatoPrec.Cells[2, 1]);
                    condGreater.Interior.ColorIndex = Struct.COLORE_VARIAZIONE_POSITIVA;

                    Excel.FormatCondition condLess = _ws.Range[rng.Rows[2, rng.Rows.Count - 1].Columns[1, 4].ToString()].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: "=" + rng.Cells[2, 1] + " < '" + mercatoPrec + "'!" + rngMercatoPrec.Cells[2, 1]);
                    condLess.Interior.ColorIndex = Struct.COLORE_VARIAZIONE_NEGATIVA;
                }
                

            }
        }
        /// <summary>
        /// Aggiorna i dati.
        /// </summary>
        public override void UpdateData()
        {
            SplashScreen.UpdateStatus("Aggiorno informazioni");
            
            CancellaDati();
            AggiornaDateTitoli();
            CaricaInformazioni();            
        }
        /// <summary>
        /// Cancella i dati.
        /// </summary>
        private void CancellaDati()
        {
            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "Gerarchia = '' OR Gerarchia IS NULL AND IdApplicazione = " + _appID;

            foreach (DataRowView entita in categoriaEntita)
            {
                List<DataRow> entitaRif =
                   (from r in categoriaEntita.Table.AsEnumerable()
                    where r["IdApplicazione"].Equals(_appID) && r["Gerarchia"].Equals(entita["SiglaEntita"])
                    select r).ToList();

                int numEntita = Math.Max(entitaRif.Count, 1);

                for (int i = 0; i < numEntita; i++)
                {
                    Range rng = new Range(_definedNames.GetRowByName(entita["SiglaEntita"], "UM", "T") + 2, _definedNames.GetColFromName("RIF" + (i + 1), "PROGRAMMAQ1"), Date.GetOreGiorno(Workbook.DataAttiva) + 2, 4);

                    _ws.Range[rng.ToString()].Value = null;
                }
            }
        }
        /// <summary>
        /// Aggiorna le date.
        /// </summary>
        public override void AggiornaDateTitoli()
        {
            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "Gerarchia = '' OR Gerarchia IS NULL AND IdApplicazione = " + _appID;

            foreach (DataRowView entita in categoriaEntita)
            {
                Range rng = new Range(_definedNames.GetRowByName(entita["SiglaEntita"], "DATA"), _struttura.colBlock + 1);
                _ws.Range[rng.ToString()].Value = Workbook.DataAttiva;
            }
        }
        /// <summary>
        /// Non ci sono grafici.
        /// </summary>
        public override void AggiornaGrafici()
        {
        }
        /// <summary>
        /// Non è necessario definire personalizzazioni.
        /// </summary>
        /// <param name="siglaEntita"></param>
        protected override void InsertPersonalizzazioni(object siglaEntita)
        {            
        }
        /// <summary>
        /// Carica le informazioni.
        /// </summary>
        public override void CaricaInformazioni()
        {
            try
            {
                if (DataBase.OpenConnection())
                {
                    SplashScreen.UpdateStatus("Carico informazioni dal DB per " + _mercato);
                    _dataInizio = Workbook.DataAttiva;

                    DataView datiApplicazione = (DataBase.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_EXPORT, "@IdApplicazione=" + _appID + ";@SiglaEntita=ALL;@SiglaCategoria=ALL;@DateFrom=" + _dataInizio.ToString("yyyyMMdd") + ";@DateTo=" + _dataInizio.ToString("yyyyMMdd")) ?? new DataTable()).DefaultView;

                    var listaEntitaInfo =
                        (from DataRowView r in datiApplicazione
                         group r by new { SiglaEntita = r["SiglaEntita"], SiglaInformazione = r["SiglaInformazione"], Riferimento = r["Riferimento"] } into g
                         select new { SiglaEntita = g.Key.SiglaEntita.ToString(), SiglaInformazione = g.Key.SiglaInformazione.ToString(), Riferimento = g.Key.Riferimento.ToString() }).ToList();

                    foreach (var entitaInfo in listaEntitaInfo)
                    {
                        SplashScreen.UpdateStatus("Scrivo informazioni " + entitaInfo.SiglaEntita);
                        datiApplicazione.RowFilter = "SiglaEntita = '" + entitaInfo.SiglaEntita + "' AND SiglaInformazione = '" + entitaInfo.SiglaInformazione + "' AND Riferimento = " + entitaInfo.Riferimento;

                        string quarter = Regex.Match(entitaInfo.SiglaInformazione, @"Q\d").Value;
                        quarter = quarter == "" ? "Q1" : quarter;

                        Range rng = new Range(_definedNames.GetRowByName(entitaInfo.SiglaEntita, "UM", "T") + 2, _definedNames.GetColFromName("RIF" + datiApplicazione[0]["Riferimento"], "PROGRAMMA" + quarter)).Extend(rowOffset: datiApplicazione.Count);

                        for (int i = 0; i <rng.Rows.Count; i++)
                            _ws.Range[rng.Rows[i].ToString()].Value = datiApplicazione[i]["Valore"];
                    }
                }
            }
            catch (Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "CaricaInformazioni InvioProgrammi SheetExport: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        protected override void InsertGrafici()
        {
        }

        public override void MakeCellsDisabled()
        {
        }
        #endregion
    }
}
