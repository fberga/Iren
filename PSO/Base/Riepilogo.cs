using System;
using System.Data;
using System.Globalization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.PSO.Base
{
    /// <summary>
    /// Interfaccia con i metodi astratti o virtuali di creazione di un foglio contenente il riepilogo.
    /// </summary>
    public abstract class ARiepilogo
    {
        #region Variabili

        protected Struct _struttura;
        protected DataView _azioni = new DataView(Workbook.Repository[DataBase.TAB.AZIONE]);
        protected DataView _categorie = new DataView(Workbook.Repository[DataBase.TAB.CATEGORIA]);
        protected DataView _entita = new DataView(Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA]);
        protected DataView _entitaAzioni = new DataView(Workbook.Repository[DataBase.TAB.ENTITA_AZIONE]);

        #endregion

        #region Metodi
        /// <summary>
        /// In un ciclo che avanza di giorno in giorno a partire da Workbook.DataAttiva per il numero di giorni definito per l'entità, esegui il delegato callback che definisce una routine specifica.
        /// </summary>
        /// <param name="callback">Delegato eseguito come corpo del ciclo.</param>
        protected void CicloGiorni(Action<int, string, DateTime> callback)
        {
            DateTime dataInizio = Workbook.DataAttiva;
            DateTime dataFine = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni);
            CicloGiorni(dataInizio, dataFine, callback);
        }
        /// <summary>
        /// In un ciclo che avanza di giorno in giorno da dataInizio a dataFine, esegui il delegato callback che definisce una routine specifica.
        /// </summary>
        /// <param name="dataInizio">Data di inizio del ciclo.</param>
        /// <param name="dataFine">Data di fine del ciclo.</param>
        /// <param name="callback">Delegato eseguito come corpo del ciclo.</param>
        protected void CicloGiorni(DateTime dataInizio, DateTime dataFine, Action<int, string, DateTime> callback)
        {
            for (DateTime giorno = dataInizio; giorno <= dataFine; giorno = giorno.AddDays(1))
            {
                int oreGiorno = Date.GetOreGiorno(giorno);
                string suffissoData = Date.GetSuffissoData(dataInizio, giorno);

                if (giorno == dataInizio && _struttura.visData0H24)
                {
                    oreGiorno++;
                }

                callback(oreGiorno, suffissoData, giorno);
            }
        }
        /// <summary>
        /// Metodo di inizializzazione dei label. Se si vuole cambiare la posizione o nascondere uno dei label è necessario eseguire l'override di questo metodo.
        /// </summary>
        public abstract void InitLabels();
        /// <summary>
        /// Launcher per il caricamento della struttura.
        /// </summary>
        public abstract void LoadStructure();
        /// <summary>
        /// Launcher per la compilazione del riepilogo in seguito allo svolgimento di un'azione. Applica di default la data attiva.
        /// </summary>
        /// <param name="siglaEntita">Entità per individuare la riga in cui scrivere.</param>
        /// <param name="siglaAzione">Azione per individuare la colonna in cui scrivere.</param>
        /// <param name="presente">Se l'azione ha portato a risultati oppure no.</param>
        public void AggiornaRiepilogo(object siglaEntita, object siglaAzione, bool presente)
        {
            AggiornaRiepilogo(siglaEntita, siglaAzione, presente, Workbook.DataAttiva);
        }
        /// <summary>
        /// Launcher per la compilazione del riepilogo in seguito allo svolgimento di un'azione.
        /// </summary>
        /// <param name="siglaEntita">Entità per individuare la riga in cui scrivere.</param>
        /// <param name="siglaAzione">Azione per individuare la colonna in cui scrivere.</param>
        /// <param name="presente">Se l'azione ha portato a risultati oppure no.</param>
        /// <param name="dataRif">La data in cui andare a scrivere. Assieme all'azione indica la colonna.</param>
        public abstract void AggiornaRiepilogo(object siglaEntita, object siglaAzione, bool presente, DateTime dataRif);
        /// <summary>
        /// Launcher per la funzione di aggiornamento dei dati del riepilogo.
        /// </summary>
        public abstract void UpdateData();

        #endregion
    }
    /// <summary>
    /// Classe base con i metodi per la creazione di un foglio contenente il riepilogo.
    /// </summary>
    public class Riepilogo : ARiepilogo
    {
        #region Variabili

        protected Excel.Worksheet _ws;
        protected DefinedNames _definedNames;
        protected int _rigaAttiva;
        protected int _colonnaInizio;
        protected int _nAzioni;
        protected static bool _resizeFatto = false;

        #endregion

        #region Costruttori

        public Riepilogo() : this(Workbook.Main)  { }

        public Riepilogo(Excel.Worksheet ws)
        {
            _ws = ws;

            _struttura = new Struct();
            _struttura.rigaBlock = 5;
            _struttura.colBlock = 59;
            try
            {
                _definedNames = new DefinedNames(_ws.Name);
            }
            catch
            {

            }
        }

        #endregion

        #region Metodi
        /// <summary>
        /// Launcher per il caricamento della struttura del riepilogo.
        /// </summary>
        public override void LoadStructure()
        {
            _colonnaInizio = _struttura.colRecap;
            _rigaAttiva = _struttura.rowRecap;

            InitLabels();
            Clear();

            if (Struct.visualizzaRiepilogo)
            {
                _categorie.RowFilter = "Operativa = 1 AND IdApplicazione = " + Workbook.IdApplicazione;
                _azioni.RowFilter = "Visibile = 1 AND Operativa = 1 AND IdApplicazione = " + Workbook.IdApplicazione;
                _entita.RowFilter = "IdApplicazione = " + Workbook.IdApplicazione;

                CreaNomiCelle();
                InitBarraTitolo();
                _rigaAttiva += 3;
                FormattaAllDati();
                InitBarraEntita();
                AbilitaAzioni();
                CaricaDatiRiepilogo();

                //Se sono in multiscreen lascio il riepilogo alla fine, altrimenti lo riporto all'inizio
                if (Screen.AllScreens.Length == 1)
                {
                    _ws.Application.ActiveWindow.SmallScroll(Type.Missing, Type.Missing, _struttura.colRecap - _struttura.colBlock - 1);
                }
                //Workbook.ScreenUpdating = false;
            }

        }
        /// <summary>
        /// Inizializza i label con dimensioni e colori caricati dal DB.
        /// </summary>
        public override void InitLabels()
        {
            _ws.Shapes.Item("lbTitolo").TextFrame.Characters().Text = Simboli.NomeApplicazione;
            _ws.Shapes.Item("lbDataInizio").TextFrame.Characters().Text = Workbook.DataAttiva.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbDataFine").TextFrame.Characters().Text = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbVersione").TextFrame.Characters().Text = "Foglio v." + Workbook.WorkbookVersion.ToString() + " Base v." + Workbook.BaseVersion.ToString() + " Core v." + Workbook.CoreVersion.ToString();
            _ws.Shapes.Item("lbUtente").TextFrame.Characters().Text = "Utente: " + Workbook.NomeUtente;

            _ws.Shapes.Item("lbTitolo").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbLinee[0], Simboli.rgbLinee[1], Simboli.rgbLinee[2]));
            _ws.Shapes.Item("lbTitolo").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbTitolo[0], Simboli.rgbTitolo[1], Simboli.rgbTitolo[2]));
            _ws.Shapes.Item("sfondo").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbLinee[0], Simboli.rgbLinee[1], Simboli.rgbLinee[2]));
            _ws.Shapes.Item("sfondo").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbSfondo[0], Simboli.rgbSfondo[1], Simboli.rgbSfondo[2]));
            _ws.Shapes.Item("lbDataInizio").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbTitolo[0], Simboli.rgbTitolo[1], Simboli.rgbTitolo[2]));
            _ws.Shapes.Item("lbDataFine").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbTitolo[0], Simboli.rgbTitolo[1], Simboli.rgbTitolo[2]));
            _ws.Shapes.Item("lbMercato").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbTitolo[0], Simboli.rgbTitolo[1], Simboli.rgbTitolo[2]));
            
            //aggiorna la scritta di modifica dati
            Simboli.ModificaDati = false;

            //aggiorna la scritta e il colore del label che mostra l'ambiente
            Handler.ChangeAmbiente(Workbook.Ambiente);

            if (Struct.intervalloGiorni > 0)
            {
                _ws.Shapes.Item("lbDataInizio").LockAspectRatio = Office.MsoTriState.msoFalse;
                _ws.Shapes.Item("lbDataInizio").Width = 26 * (float)_ws.Columns[1].Width;
                _ws.Shapes.Item("lbDataFine").Visible = Office.MsoTriState.msoTrue;
                _ws.Shapes.Item("lbDataInizio").LockAspectRatio = Office.MsoTriState.msoTrue;
            }
            else
            {
                _ws.Shapes.Item("lbDataInizio").LockAspectRatio = Office.MsoTriState.msoFalse;
                _ws.Shapes.Item("lbDataInizio").Width = 54 * (float)_ws.Columns[1].Width;
                _ws.Shapes.Item("lbDataFine").Visible = Office.MsoTriState.msoFalse;
                _ws.Shapes.Item("lbDataInizio").LockAspectRatio = Office.MsoTriState.msoTrue;
            }
            Workbook.ScreenUpdating = false;
        }
        /// <summary>
        /// Metodo per eliminare la struttura esistente dal foglio e prepararlo alla nuova che verrà caricata.
        /// </summary>
        protected virtual void Clear()
        {
            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            _ws.UsedRange.EntireColumn.Delete();
            _ws.UsedRange.FormatConditions.Delete();
            _ws.UsedRange.Interior.ColorIndex = 2;
            _ws.UsedRange.EntireRow.Hidden = false;
            _ws.UsedRange.Font.Size = 8;
            _ws.UsedRange.NumberFormat = "General";
            _ws.UsedRange.Font.Name = "Verdana";

            _ws.Columns.ColumnWidth = 9;

            _ws.Range[Range.GetRange(1, 1, 1, _struttura.colRecap - 1)].EntireColumn.ColumnWidth = Struct.cell.width.empty;            
            _ws.Rows[1].RowHeight = Struct.cell.height.empty;

            ((Excel._Worksheet)_ws).Activate();
            _ws.Application.ActiveWindow.FreezePanes = false;
            _ws.Cells[_struttura.rigaBlock, _struttura.colBlock].Select();
            _ws.Application.ActiveWindow.ScrollColumn = 1;
            _ws.Application.ActiveWindow.ScrollRow = 1;
            _ws.Application.ActiveWindow.FreezePanes = true;
            Workbook.ScreenUpdating = false;
        }
        /// <summary>
        /// Crea la struttura dei nomi del riepilogo definendo 3 righe di titolo (DATA, AZIONE PARDE, AZIONE), una riga per ogni entità, e una colonna per ogni AZIONE con la DATA di riferimento.
        /// </summary>
        protected virtual void CreaNomiCelle()
        {
            //inserisco tutte le righe
            _definedNames.AddName(_rigaAttiva++, "DATA");
            _definedNames.AddName(_rigaAttiva++, "AZIONI_PADRE");
            _definedNames.AddName(_rigaAttiva++, "AZIONI");

            foreach (DataRowView categoria in _categorie)
            {
                if(Workbook.Repository.Applicazione["VisCategoriaRiepilogo"].Equals("1"))
                    _definedNames.AddName(_rigaAttiva++, categoria["SiglaCategoria"]);

                _entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                foreach (DataRowView e in _entita)
                {
                    _definedNames.AddName(_rigaAttiva, e["SiglaEntita"]);
                    _definedNames.AddGOTO(e["SiglaEntita"], Range.R1C1toA1(_rigaAttiva++, _colonnaInizio));
                }
            }
            
            //inserisco tutte le colonne
            _definedNames.AddCol(_colonnaInizio++, "COLONNA_ENTITA");
            CicloGiorni((oreGiorno, suffissoData, giorno) => 
            {
                foreach (DataRowView azione in _azioni)
                {
                    if (azione["Gerarchia"] != DBNull.Value)
                        _definedNames.AddCol(_colonnaInizio++, azione["SiglaAzione"], suffissoData);
                }
            });
            _definedNames.DumpToDataSet();
        }
        /// <summary>
        /// Aggiorna la colorazione della barra superiore del riepilogo in base allo schema colori basato sui giorni.
        /// </summary>
        protected void UpdateDayColor()
        {
            DataView azioni = new DataView(Workbook.Repository[DataBase.TAB.AZIONE]);
            azioni.RowFilter = "Visibile = 1 AND Operativa = 1 AND IdApplicazione = " + Workbook.IdApplicazione;

            Range rngTitleBar = new Range(_definedNames.GetFirstRow(), _definedNames.GetFirstCol() + 1, 3, azioni.Count);

            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                ASheet.AssegnaColori(_ws.Range[rngTitleBar.ToString()], giorno);

                rngTitleBar.StartColumn += azioni.Count;
            });
        }
        /// <summary>
        /// Inizializza la barra del titolo.
        /// </summary>
        protected void InitBarraTitolo()
        {
            Range rngTitleBar = new Range(_definedNames.GetFirstRow(), _definedNames.GetFirstCol() + 1, 3, _azioni.Count);
            Range rngData = rngTitleBar.Cells[0, 0];
            Range rngAzioniPadre = rngTitleBar.Cells[1, 0];
            Range rngAzioni = rngTitleBar.Cells[2, 0];

            string azionePadre = "";
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                rngTitleBar.StartColumn = rngAzioni.StartColumn;
                _ws.Range[rngTitleBar.ToString()].Style = "Barra titolo riepilogo";
                _ws.Range[rngTitleBar.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

                foreach (DataRowView azione in _azioni)
                {
                    if (!azione["Gerarchia"].Equals(azionePadre))
                    {
                        rngAzioniPadre.ColOffset = rngAzioni.StartColumn - rngAzioniPadre.StartColumn;
                        Style.RangeStyle(_ws.Range[rngAzioniPadre.ToString()], merge:true, fontSize:9);
                        _ws.Range[rngAzioniPadre.ToString()].Value = azionePadre;
                        azionePadre = azione["Gerarchia"].ToString();
                        rngAzioniPadre.StartColumn = rngAzioni.StartColumn;
                    }
                    _ws.Range[rngAzioni.ToString()].Value = azione["DesAzioneBreve"];
                    Style.RangeStyle(_ws.Range[rngAzioni.ToString()], fontSize:7);
                    rngAzioni.StartColumn++;
                }
                rngAzioniPadre.ColOffset = rngAzioni.StartColumn - rngAzioniPadre.StartColumn;
                Style.RangeStyle(_ws.Range[rngAzioniPadre.ToString()], merge:true, fontSize:9);
                _ws.Range[rngAzioniPadre.ToString()].Value = azionePadre;
                azionePadre = "";
                rngAzioniPadre.StartColumn = rngAzioni.StartColumn;

                rngData.ColOffset = rngAzioni.StartColumn - rngData.StartColumn;
                Style.RangeStyle(_ws.Range[rngData.ToString()], merge:true, fontSize:10, numberFormat:"ddd d mmm yyyy");
                _ws.Range[rngData.ToString()].Value = giorno;
                rngData.StartColumn = rngAzioni.StartColumn;
            });

            UpdateDayColor();
        }
        /// <summary>
        /// Formatta il range che conterrà tutti i dati del riepilogo. Le celle sono tutte disabilitate e verranno abilitate nella funzione AbilitaAzioni.
        /// </summary>
        protected void FormattaAllDati()
        {
            Range rngAll = new Range(_definedNames.GetFirstRow(), _definedNames.GetFirstCol() + 1, _definedNames.GetRowOffset(), _definedNames.GetColOffsetRiepilogo() - 1);
            Range rngData = new Range(_definedNames.GetFirstRow() + 3, _definedNames.GetFirstCol(), _definedNames.GetRowOffset() - 3, _definedNames.GetColOffsetRiepilogo());
            
            _ws.Range[rngData.ToString()].Style = "Area dati riepilogo";
            _ws.Range[rngData.Columns[0].ToString()].Style = "Lista entita riepilogo";
            _ws.Range[rngData.Columns[0].ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            Excel.Range xlrng = _ws.Range[rngAll.Rows[1, rngAll.Rows.Count - 1].ToString()];
            //trovo tutte le aree unite e creo il blocco col bordo grosso
            int i = 0;
            int colspan = 0;
            while (i < xlrng.Columns.Count)
            {
                colspan = xlrng.Cells[1, i + 1].MergeArea().Columns.Count;
                _ws.Range[rngAll.Columns[i, i + colspan - 1].ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
                _ws.Range[rngAll.Columns[i, i + colspan - 1].ToString()].Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                i += colspan;
            }

            
            _ws.Range[rngAll.ToString()].EntireColumn.AutoFit();
            if(rngAll.ColOffset > 1)
            {
                //calcolo la massima dimensione delle colonne e la riapplico a tutto il riepilogo
                double maxWidth = double.MinValue;
                foreach (Range col in rngAll.Columns)
                    maxWidth = Math.Max(_ws.Range[col.ToString()].ColumnWidth, maxWidth);

                foreach (Range col in rngAll.Columns)
                    _ws.Range[col.ToString()].ColumnWidth = maxWidth + 1;
            }
        }
        /// <summary>
        /// Crea la barra laterale con la lista di tutte le entità.
        /// </summary>
        protected void InitBarraEntita()
        {
            int row = _definedNames.GetFirstRow() + 3;
            foreach (DataRowView categoria in _categorie)
            {
                Range rng = new Range(row, _definedNames.GetFirstCol(), 1, _definedNames.GetColOffsetRiepilogo());

                if (Workbook.Repository.Applicazione["VisCategoriaRiepilogo"].Equals("1"))
                {
                    Style.RangeStyle(_ws.Range[rng.ToString()], style: "Lista categorie riepilogo", borders: "[left:medium,top:medium,right:medium]", merge: true);
                    _ws.Range[rng.Columns[0].ToString()].Value = categoria["DesCategoria"];
                    rng.StartRow++;
                }

                _entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                foreach (DataRowView entita in _entita)
                {
                    _ws.Range[rng.Columns[0].ToString()].Value = (entita["Gerarchia"] is DBNull ? "" : "     ") + entita["DesEntita"];
                    _ws.Range[rng.Columns[0].ToString()].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                    rng.StartRow++;
                }
                row = rng.StartRow;
            }

            if (Workbook.Repository.Applicazione["VisCategoriaRiepilogo"].Equals("0"))
            {
                Range firstRow = new Range(_definedNames.GetFirstRow() + 3, _definedNames.GetFirstCol(), 1, _definedNames.GetColOffsetRiepilogo());
                _ws.Range[firstRow.ToString()].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
            }

            _ws.Columns[_struttura.colRecap].EntireColumn.AutoFit();
        }
        /// <summary>
        /// Abilita le azioni per ogni entità.
        /// </summary>
        protected void AbilitaAzioni()
        {
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                foreach (DataRowView azione in _entitaAzioni)
                {


                    if (azione["Giorno"] is DBNull || azione["Giorno"].ToString().Contains(suffissoData))
                    {
                        Range cellaAzione = new Range(_definedNames.GetRowByName(azione["SiglaEntita"]), _definedNames.GetColFromName(azione["SiglaAzione"], suffissoData));
                        _ws.Range[cellaAzione.ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
                        _ws.Range[cellaAzione.ToString()].Interior.ColorIndex = 2;
                    }
                }
            });
        }
        /// <summary>
        /// Carico i dati e i commenti che devono essere scritti nelle celle.
        /// </summary>
        protected void CaricaDatiRiepilogo()
        {
            try
            {
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    if (DataBase.OpenConnection())
                    {
                        DataView datiRiepilogo = (DataBase.Select(DataBase.SP.APPLICAZIONE_RIEPILOGO, "@Data=" + giorno.ToString("yyyyMMdd")) ?? new DataTable()).DefaultView;
                        foreach (DataRowView valore in datiRiepilogo)
                        {
                            Range cellaAzione = new Range(_definedNames.GetRowByName(valore["SiglaEntita"]), _definedNames.GetColFromName(valore["SiglaAzione"], suffissoData));

                            Excel.Range rng = _ws.Range[cellaAzione.ToString()];

                            if (valore["Presente"].Equals("1"))
                            {
                                rng.ClearComments();
                                DateTime data = DateTime.ParseExact(valore["Data"].ToString(), "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                                rng.AddComment("Utente: " + valore["Utente"] + "\nData: " + data.ToString("dd MMM yyyy") + "\nOra: " + data.ToString("HH:mm"));
                                rng.Value = "OK";
                                Style.RangeStyle(rng, foreColor: 1, bold: true, fontSize: 9, backColor: 4, align: Excel.XlHAlign.xlHAlignCenter);
                            }
                            else
                            {
                                rng.ClearComments();
                                rng.Value = "Non presente";
                                Style.RangeStyle(rng, foreColor: 3, bold: false, fontSize: 7, backColor: 2, align: Excel.XlHAlign.xlHAlignCenter);
                            }
                        }
                    }
                });
            }
            catch (Exception e)
            {
                Workbook.InsertLog(PSO.Core.DataBase.TipologiaLOG.LogErrore, "CaricaDatiRiepilogo: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Launcher per la compilazione del riepilogo in seguito allo svolgimento di un'azione.
        /// </summary>
        /// <param name="siglaEntita">Entità per individuare la riga in cui scrivere.</param>
        /// <param name="siglaAzione">Azione per individuare la colonna in cui scrivere.</param>
        /// <param name="presente">Se l'azione ha portato a risultati oppure no.</param>
        /// <param name="dataRif">La data in cui andare a scrivere. Assieme all'azione indica la colonna.</param>
        public override void AggiornaRiepilogo(object siglaEntita, object siglaAzione, bool presente, DateTime dataRif)
        {

            if (dataRif - Workbook.DataAttiva <= new TimeSpan(Struct.intervalloGiorni, 0, 0, 0))
            {
                if (Struct.visualizzaRiepilogo && !Simboli.EmergenzaForzata)
                {
                    Range cell = _definedNames.Get(siglaEntita, siglaAzione, Date.GetSuffissoData(dataRif));
                    Excel.Range rng = _ws.Range[cell.ToString()];
                    if (presente)
                    {
                        string commento = "Utente: " + Workbook.NomeUtente + "\nData: " + DateTime.Now.ToString("dd MMM yyyy") + "\nOra: " + DateTime.Now.ToString("HH:mm");
                        rng.ClearComments();
                        rng.AddComment(commento).Visible = false;
                        rng.Value = "OK";
                        Style.RangeStyle(rng, foreColor: 1, bold: true, fontSize: 9, backColor: 4, align: Excel.XlHAlign.xlHAlignCenter);
                    }
                    else
                    {
                        rng.ClearComments();
                        rng.Value = "Non presente";
                        Style.RangeStyle(rng, foreColor: 3, bold: false, fontSize: 7, backColor: 2, align: Excel.XlHAlign.xlHAlignCenter);
                    }
                }
            }
        }
        /// <summary>
        /// Cancella tutti i dati contenuti nel riepilogo
        /// </summary>
        private void CancellaDati()
        {
            Range rngData = new Range(_definedNames.GetFirstRow() + 3, _definedNames.GetFirstCol() + 1, _definedNames.GetRowOffset() - 3, _definedNames.GetColOffsetRiepilogo() - 1);
            _ws.Range[rngData.ToString()].Value = null;
            _ws.Range[rngData.ToString()].Interior.ColorIndex = 2;
            _ws.Range[rngData.ToString()].ClearComments();
        }
        /// <summary>
        /// Aggiorna le date nei titoli e label del riepilogo.
        /// </summary>
        protected void AggiornaDate()
        {
            _ws.Shapes.Item("lbDataInizio").TextFrame.Characters().Text = Workbook.DataAttiva.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbDataFine").TextFrame.Characters().Text = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("ddd d MMM yyyy");
            
            if (Struct.visualizzaRiepilogo)
            {
                _azioni.RowFilter = "Visibile = 1 AND Operativa = 1 AND Gerarchia IS NOT NULL AND IdApplicazione = " + Workbook.IdApplicazione;

                if (_azioni.Count > 0)
                {
                    CicloGiorni((oreGiorno, suffissoData, giorno) =>
                    {
                        Range cell = new Range(_definedNames.GetRowByName("DATA"), _definedNames.GetColFromName(_azioni[0]["SiglaAzione"], suffissoData));
                        _ws.Range[cell.ToString()].Value = giorno;
                    });
                }
                
                _azioni.RowFilter = "Visibile = 1 AND Operativa = 1 AND IdApplicazione = " + Workbook.IdApplicazione;

                //ridimensiono le celle per adattarle ad eventuali modifiche nei contenuti
                //Range rngAll = new Range(_definedNames.GetFirstRow(), _definedNames.GetFirstCol() + 1, _definedNames.GetRowOffset(), _definedNames.GetColOffsetRiepilogo() - 1);
                //_ws.Range[rngAll.ToString()].EntireColumn.AutoFit();
                //if (rngAll.ColOffset > 1)
                //{
                //    //calcolo la massima dimensione delle colonne e la riapplico a tutto il riepilogo
                //    double maxWidth = double.MinValue;
                //    foreach (Range col in rngAll.Columns)
                //        maxWidth = Math.Max(_ws.Range[col.ToString()].ColumnWidth, maxWidth);

                //    foreach (Range col in rngAll.Columns)
                //        _ws.Range[col.ToString()].ColumnWidth = maxWidth;
                //}

                UpdateDayColor();
            }
        }
        /// <summary>
        /// Launcher per la funzione di aggiornamento dei dati del riepilogo.
        /// </summary>
        public override void UpdateData()
        {
            if (_definedNames != null)
            {
                AggiornaDate();

                if (Struct.visualizzaRiepilogo)
                {
                    CancellaDati();
                    AbilitaAzioni();
                    CaricaDatiRiepilogo();
                }
            }
        }
        /// <summary>
        /// In emergenza, permette di impostare lo stile delle celle a disabilitato (infatti il riepilogo non verrebbe aggiornato in ogni caso).
        /// </summary>
        private void DisabilitaTutto()
        {
            Range rngData = new Range(_definedNames.GetFirstRow() + 3, _definedNames.GetFirstCol() + 1, _definedNames.GetRowOffset() - 3, _definedNames.GetColOffsetRiepilogo() - 1);

            Style.RangeStyle(_ws.Range[rngData.ToString()], pattern: Excel.XlPattern.xlPatternCrissCross);
        }
        /// <summary>
        /// Mette il riepilogo in stato di emergenza.
        /// </summary>
        public void RiepilogoInEmergenza()
        {
            if (Struct.visualizzaRiepilogo)
            {
                AggiornaDate();
                DisabilitaTutto();
            }
        }

        #endregion
    }
}
