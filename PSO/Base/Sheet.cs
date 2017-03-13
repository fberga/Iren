using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Base
{
    /// <summary>
    /// Interfaccia con i metodi astratti o virtuali di creazione di un foglio contenente dati riferiti a impianti.
    /// </summary>
    public abstract class ASheet
    {
        #region Variabili

        protected Struct _struttura;
        protected DateTime _dataInizio;
        protected DateTime _dataFine;
        protected int _visSelezione;

        protected static bool _protetto = true;

        #endregion

        #region Metodi

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
                string suffissoData = Date.GetSuffissoData(_dataInizio, giorno);

                if (Struct.tipoVisualizzazione == "V" || Struct.tipoVisualizzazione == "R")
                {
                    oreGiorno = 25;
                    suffissoData = Date.GetSuffissoData(Workbook.DataAttiva, giorno);
                }
                callback(oreGiorno, suffissoData, giorno);
            }
        }
        /// <summary>
        /// In un ciclo che avanza di giorno in giorno a partire da Workbook.DataAttiva per il numero di giorni definito per l'entità, esegui il delegato callback che definisce una routine specifica.
        /// </summary>
        /// <param name="callback">Delegato eseguito come corpo del ciclo.</param>
        protected void CicloGiorni(Action<int, string, DateTime> callback)
        {
            CicloGiorni(_dataInizio, _dataFine, callback);
        }
        /// <summary>
        /// Metodo di caricamento della struttura del foglio.
        /// </summary>
        public abstract void LoadStructure();
        /// <summary>
        /// Metodo di aggiornamento dei dati del foglio.
        /// </summary>
        public abstract void UpdateData();
        /// <summary>
        /// Metodo di aggiornamento delle date dei titolo.
        /// </summary>
        public abstract void AggiornaDateTitoli();
        /// <summary>
        /// Metodo di inserimento dei grafici
        /// </summary>
        protected abstract void InsertGrafici();
        /// <summary>
        /// Metodo di aggiornamento dei grafici.
        /// </summary>
        public abstract void AggiornaGrafici();
        /// <summary>
        /// Metodo che permette di aggiungere delle customizzazioni durante la creazione della struttura.
        /// </summary>
        /// <param name="siglaEntita"></param>
        protected virtual void InsertPersonalizzazioni(object siglaEntita) { }
        /// <summary>
        /// Metodo per il caricamento delle informazioni.
        /// </summary>
        public abstract void CaricaInformazioni();

        public abstract void MakeCellsDisabled();
        
        #endregion

        #region Proprietà Statiche

        /// <summary>
        /// Restituisce o imposta la proprietà di protezione del workbook e dei fogli in esso contenuti.
        /// </summary>
        public static bool Protected
        {
            get { return _protetto; }
            set
            {
                if (_protetto != value)
                {
                    if (value)
                        Workbook.WB.Protect(Workbook.Password);
                    else
                        Workbook.WB.Unprotect(Workbook.Password);

                    _protetto = value;

                    foreach (Excel.Worksheet ws in Workbook.Sheets)
                    {
                        if (value)
                            if (ws.Name == "Log")
                                ws.Protect(Workbook.Password, AllowSorting: true, AllowFiltering: true);
                            else
                                ws.Protect(Workbook.Password);
                        else
                            ws.Unprotect(Workbook.Password);

                        Marshal.ReleaseComObject(ws);
                    }
                }
            }
        }

        #endregion

        #region Metodi Statici

        /// <summary>
        /// Metodo per abilitare la modifica nelle informazioni per cui è concessa la modifica da DB.
        /// </summary>
        /// <param name="abilita">La modifica è abilitata se la proprietà è a true.</param>
        public static void AbilitaModifica(bool abilita)
        {
            DataView categorie = Workbook.Repository[DataBase.TAB.CATEGORIA].DefaultView;
            categorie.RowFilter = "Operativa = '1' AND IdApplicazione = " + Workbook.IdApplicazione;
            
            bool prot = Protected;
            if(prot)
                Protected = false;

            foreach (DataRowView categoria in categorie)
            {
                Excel.Worksheet ws = Workbook.Sheets[categoria["DesCategoria"].ToString()];
                DefinedNames definedNames = new DefinedNames(categoria["DesCategoria"].ToString(), DefinedNames.InitType.Editable);

                foreach (string range in definedNames.Editable.Values)
                {
                    string[] subRanges = range.Split(',');
                    if (subRanges.Length == 1 && ws.Range[subRanges[0]].Cells.Count == 1)
                    {
                        ws.Range[subRanges[0]].Locked = !abilita;
                    }
                    else if (ws.Range[subRanges[0]].EntireRow.Hidden == false)
                    {
                        foreach (string subRange in subRanges)
                        {
                            Range rng = new Range(subRange);
                            //sono in un range dei dati
                            if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1") && rng.Columns.Count > 22 && rng.Columns.Count < 26)
                            {
                                rng = GetModifiableRange(DateTime.Now.Hour, rng);
                            }
                            else if (Workbook.IdApplicazione == 18)
                            {
                                int hour = Simboli.GetMarketOffsetMI(Workbook.Mercato, Workbook.DataAttiva) - 1;
                                rng.StartColumn += hour;
                                rng.ColOffset -= hour;
                            }
                            ws.Range[rng.ToString()].Locked = !abilita;
                        }
                    }
                }
            }
            
            if(prot)
                Protected = true;
        }
        /// <summary>
        /// Restituisce il range che è effettivamente modificabile in caso ci siano mercati chiusi. Presuppone che la riga considerata sia una riga di dati.
        /// </summary>
        /// <param name="hour">Ora in cui effettuare il controllo.</param>
        /// <param name="rng">Range originale.</param>
        /// <returns>Ritorna un range più piccolo dell'originale che contiene le ore effettivamente modificabili.</returns>
        public static Range GetModifiableRange(int hour, Range rng)
        {
            int startHour = Simboli.GetMarketOffset(hour);
            Range o = new Range(rng);

            o.StartColumn += startHour;
            o.ColOffset -= startHour;

            return o;
        }
        /// <summary>
        /// Metodo che registra in DataBase.LocalDB le modifiche effettuate dall'utente durante il periodo in cui la modifica è attiva. Lavora con le sole entità modificate in modo da non sovraccaricare di modifiche il DB.
        /// </summary>
        public static void SalvaModifiche()
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

                    entitaInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND ((FormulaInCella = '1' AND WB = '0' AND SalvaDB = '1') OR (WB <> '0' AND SalvaDB = '1')) AND IdApplicazione = " + Workbook.IdApplicazione;

                    DataTable entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA];
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
                            Handler.StoreEdit(ws.Range[rng.ToString()], 0, true);
                        }
                        else
                        {
                            //TODO fare ciclo giorni nel caso di visualizzazione verticale: le informazioni si dividono per i giorni e non sono in linea!!!
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Assegna il colore al range confrontando il suo contenuto con lo schema definito dagli utenti: Giallo per le date antecedenti a oggi, verde oggi, azzurro domani, arancione dopodomani, grigio il resto.
        /// </summary>
        /// <param name="rng">Range su cui applicare la colorazione.</param>
        /// <param name="giorno">Giorno con cui fare il confronto.</param>
        public static void AssegnaColori(Excel.Range rng, DateTime giorno)
        {
            Style.RangeStyle(rng, pattern: Excel.XlPattern.xlPatternNone);

            if (giorno.Date < DateTime.Now.Date)
                rng.Interior.Color = System.Drawing.Color.FromArgb(240, 230, 140);
            else if (giorno.Date == DateTime.Now.Date)
                rng.Interior.Color = System.Drawing.Color.FromArgb(144, 238, 144);
            else if (giorno.Date == DateTime.Now.Date.AddDays(1))
                rng.Interior.Color = System.Drawing.Color.FromArgb(135, 206, 250);
            else if (giorno.Date == DateTime.Now.Date.AddDays(2))
                rng.Interior.Color = System.Drawing.Color.FromArgb(244, 164, 96);
            else
                rng.Interior.Color = System.Drawing.Color.FromArgb(192, 192, 192);
        }

        #endregion
    }
    /// <summary>
    /// Classe base con i metodi per la creazione di un foglio contenente dati riferiti a impianti.
    /// </summary>
    public class Sheet : ASheet, IDisposable
    {
        #region Variabili

        protected Excel.Worksheet _ws;
        protected object _siglaCategoria;
        protected DefinedNames _definedNames;
        protected int _intervalloOre;
        protected int _rigaAttiva;
        protected bool _disposed = false;
        protected int _intervalloGiorniMax;
        protected Dictionary<object, DateTime> _dataFineUP = new Dictionary<object, DateTime>();

        #endregion

        #region Costruttori

        public Sheet(Excel.Worksheet ws)
        {
            _ws = ws;

            DataView categorie = Workbook.Repository[DataBase.TAB.CATEGORIA].DefaultView;
            categorie.RowFilter = "DesCategoria = '" + ws.Name + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            _siglaCategoria = categorie[0]["SiglaCategoria"];

            AggiornaParametriSheet();
            _definedNames = new DefinedNames(_ws.Name);

            //carico la massima datafine in maniera da creare la barra navigazione della dimensione giusta (compresa la definizione dei giorni se necessario)
            DataView entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].DefaultView;
            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND IdApplicazione = " + Workbook.IdApplicazione;            

            foreach (DataRowView entita in categoriaEntita)
            {
                entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_STRUTTURA' AND IdApplicazione = " + Workbook.IdApplicazione;
                int intervalloGiorni = entitaProprieta.Count > 0 ? int.Parse(entitaProprieta[0]["Valore"].ToString()) : Struct.intervalloGiorni;

                _dataFineUP.Add(entita["SiglaEntita"], Workbook.DataAttiva.AddDays(intervalloGiorni));
                _intervalloGiorniMax = Math.Max(_intervalloGiorniMax, intervalloGiorni);
            }
        }
        ~Sheet()
        {
            Dispose();
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Colora le celle GOTO in base alla data attiva e al giorno corrente.
        /// </summary>
        protected void ColoraGOTO()
        {
            if (Struct.tipoVisualizzazione == "V" && Struct.visLinkEntita)
            {
                List<string> gotos = _definedNames.GetAllFromAddressGOTO();

                foreach (string address in gotos)
                {
                    DateTime giorno = _ws.Range[address].Value;
                    AssegnaColori(_ws.Range[address], giorno);
                }
            }
        }
        /// <summary>
        /// Colora le intestazioni di data e ora in base alla data attiva e al giorno corrente.
        /// </summary>
        protected void ColoraDataOra()
        {
            CicloGiorni(Workbook.DataAttiva, Workbook.DataAttiva.AddDays(_intervalloGiorniMax), (oreGiorno, suffissoData, giorno) =>
            {
                if (Struct.tipoVisualizzazione != "R")
                {
                    int row = 0;
                    if (Struct.tipoVisualizzazione == "O")
                    {
                        row = _struttura.rigaBlock - 2;
                    }
                    else if (Struct.tipoVisualizzazione == "V")
                    {
                        row = _definedNames.Get(Date.GetSuffissoData(giorno), "T").StartRow;
                    }

                    Range rng = new Range(row, _definedNames.GetColFromDate(giorno), 2, oreGiorno);

                    AssegnaColori(_ws.Range[rng.ToString()], giorno);

                    if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1"))
                    {
                        int offset = Simboli.GetMarketOffset(DateTime.Now.Hour);
                        //06/02/2017 MOD: prendo il minimo tra l'ora del mercato e le ore giorno
                        Range rngDisabled = new Range(row + 1, _definedNames.GetColFromDate(giorno), 1, Math.Min(offset, oreGiorno));
                        Style.RangeStyle(_ws.Range[rngDisabled.ToString()], pattern: Excel.XlPattern.xlPatternGray50);
                    }

                    if (Struct.tipoVisualizzazione == "V")
                    {
                        //coloro titolo Verticale
                        rng = new Range(row + 2, _struttura.colBlock - _visSelezione - 1);
                        AssegnaColori(_ws.Range[rng.ToString()].MergeArea, giorno);
                    }
                }
            });
        }
        /// <summary>
        /// Legge da LocalDB i parametri dell'applicazione e reinizializza tutte le strutture del workbook e del foglio.
        /// </summary>
        protected void AggiornaParametriSheet()
        {
            _struttura = new Struct();

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;

            _struttura.rigaBlock = (int)Workbook.Repository.Applicazione["RowBlocco"];
            _struttura.rigaGoto = (int)Workbook.Repository.Applicazione["RowGoto"];
            _struttura.visData0H24 = Workbook.Repository.Applicazione["VisData0H24"].ToString() == "1";
            
            _struttura.visSelezione =
                (from c in Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].AsEnumerable()
                 join e in Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].AsEnumerable()
                 on new { SiglaEntita = c["SiglaEntita"], IdApplicazione = c["IdApplicazione"] } equals
                    new { SiglaEntita = e["SiglaEntita"], IdApplicazione = e["IdApplicazione"] }
                 where c["SiglaCategoria"].Equals(_siglaCategoria)
                    && c["IdApplicazione"].Equals(Workbook.IdApplicazione)
                    && (int)e["Selezione"] > 0
                 select e).Count() > 0;

            _struttura.colBlock = (int)Workbook.Repository.Applicazione["ColBlocco"] + (_struttura.visSelezione ? 1 : 0);
            Struct.tipoVisualizzazione = Workbook.Repository.Applicazione["TipoVisualizzazione"] is DBNull ? "O" : Workbook.Repository.Applicazione["TipoVisualizzazione"].ToString();
            Struct.intervalloGiorni = Workbook.Repository.Applicazione["IntervalloGiorniEntita"] is DBNull ? 0 : (int)Workbook.Repository.Applicazione["IntervalloGiorniEntita"];
            Struct.visualizzaRiepilogo = Workbook.Repository.Applicazione["VisRiepilogo"] is DBNull ? true : Workbook.Repository.Applicazione["VisRiepilogo"].Equals("1");
            Struct.visLinkEntita = Workbook.Repository.Applicazione["VisLinkEntita"] is DBNull ? true : Workbook.Repository.Applicazione["VisLinkEntita"].Equals("1");

            _visSelezione = (_struttura.visSelezione ? 3 : 2);

            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL) AND IdApplicazione = " + Workbook.IdApplicazione;
            if (Struct.visLinkEntita)
                _struttura.numEleMenu = (Struct.tipoVisualizzazione == "O" ? categoriaEntita.Count : (Struct.intervalloGiorni + 1));
            else
                _struttura.numEleMenu = 0;
            _struttura.numRigheMenu = 1;
            if (_struttura.numEleMenu > 8)
            {
                int tmp = _struttura.numEleMenu;
                while (tmp / 8 > 0)
                {
                    _struttura.rigaBlock++;
                    _struttura.numRigheMenu++;
                    tmp /= 8;
                }
            }
        }
        /// <summary>
        /// Launcher per l'aggiornamento della struttura. Definisce anche le colonne in base all'intervallo massimo di giorni delle entità presenti nel foglio.
        /// </summary>
        public override void LoadStructure()
        {
            SplashScreen.UpdateStatus("Aggiorno struttura " + _ws.Name);

            DataTable entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA];
            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;

            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL) AND IdApplicazione = " + Workbook.IdApplicazione;

            _dataInizio = Workbook.DataAttiva;
            _dataFine = Workbook.DataAttiva.AddDays(Struct.tipoVisualizzazione == "O" ? _intervalloGiorniMax : 0);

            //Definizione dei nomi delle colonne
            _definedNames.DefineDates(_dataInizio, _dataFine, _struttura.colBlock, _struttura.visData0H24);

            Clear();
            
            InitBarraNavigazione();
            

            _rigaAttiva = _struttura.rigaBlock + 1;

            foreach (DataRowView entita in categoriaEntita)
            {
                string siglaEntita = "" + entita["SiglaEntita"];

                if (Struct.tipoVisualizzazione == "O")
                {
                    _dataFine = _dataFineUP[siglaEntita];
                    InitBloccoEntita(entita);
                }
                else if (Struct.tipoVisualizzazione == "V")
                {
                    CicloGiorni(_dataInizio, _dataInizio.AddDays(Struct.intervalloGiorni), (oreGiorno, suffissoData, giorno) =>
                    {
                        _dataFine = _dataInizio = giorno;
                        InitBloccoEntita(entita);
                    });
                }
                else if (Struct.tipoVisualizzazione == "R")
                {
                    InitBloccoEntita(entita);
                }
            }

            ColoraDataOra();

            categoriaEntita.RowFilter = "IdApplicazione = " + Workbook.IdApplicazione;

            _definedNames.DumpToDataSet();
            CaricaInformazioni();
            AggiornaGrafici();
            MakeCellsDisabled();
        }
        /// <summary>
        /// Metodo per eliminare la struttura esistente dal foglio e prepararlo alla nuova che verrà caricata.
        /// </summary>
        protected void Clear()
        {
            SplashScreen.UpdateStatus("Cancello struttura foglio '" + _ws.Name + "'");

            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            if (_ws.ChartObjects().Count > 0)
                _ws.ChartObjects().Delete();

            _ws.Rows.ClearContents();
            _ws.Rows.ClearComments();
            _ws.Rows.FormatConditions.Delete();
            _ws.Rows.Validation.Delete();
            _ws.Rows.EntireRow.Hidden = false;
            _ws.Rows.Style = "Normal";
            _ws.Rows.UnMerge();

            _ws.Rows.RowHeight = Struct.cell.height.normal;
            _ws.Columns.ColumnWidth = Struct.cell.width.dato;

            _ws.Rows["1:" + (_struttura.rigaBlock - 1)].RowHeight = Struct.cell.height.empty;

            for (int i = 0; i < _struttura.numRigheMenu; i++)
                _ws.Rows[_struttura.rigaGoto + i].RowHeight = Struct.cell.height.normal;

            _ws.Columns[1].ColumnWidth = Struct.cell.width.empty;
            _ws.Columns[2].ColumnWidth = Struct.cell.width.entita;

            if (!Aggiorna._freezePanes.ContainsKey(_ws.Name) ||
                (Aggiorna._freezePanes[_ws.Name].Item1 != _struttura.rigaBlock || Aggiorna._freezePanes[_ws.Name].Item2 != _struttura.colBlock))
            {
                ((Excel._Worksheet)_ws).Activate();
                _ws.Application.ActiveWindow.FreezePanes = false;
                _ws.Cells[_struttura.rigaBlock, _struttura.colBlock].Select();
                _ws.Application.ActiveWindow.FreezePanes = true;
                ((Excel._Worksheet)Workbook.Main).Activate();
            }
            

            int colInfo = _struttura.colBlock - _visSelezione;
            _ws.Columns[colInfo].ColumnWidth = Struct.cell.width.informazione;
            _ws.Columns[colInfo + 1].ColumnWidth = Struct.cell.width.unitaMisura;
            if (_struttura.visSelezione)
                _ws.Columns[colInfo + 2].ColumnWidth = 2.5;
        }
        /// <summary>
        /// Inizializza la barra di navigazione nella parte alta del foglio applicandovi lo stile "Top menu GOTO". Definisce tutte le celle e genera gli oggetti GOTO per i tasti del menù e applica lo stile "Barra navigazione con nomi" se la visualizzazione è Orizzontale altrimenti "Barra navigazione con date". Se la tipologia di visualizzazione è Orizzontale, aggiunge anche la barra della data e delle ore (applicando "Barra della data").
        /// </summary>
        protected void InitBarraNavigazione()
        {
            SplashScreen.UpdateStatus("Inizializzo barra di navigazione '" + _ws.Name + "'");

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL ) AND IdApplicazione = " + Workbook.IdApplicazione;

            int dataOreTot = (Struct.tipoVisualizzazione == "O" ? Date.GetOreIntervallo(_dataInizio, _dataFine) : 25) + (_struttura.visData0H24 ? 1 : 0);
                
            Excel.Range gotoBar = _ws.Range[_ws.Cells[2, 2], _ws.Cells[_struttura.rigaGoto + _struttura.numRigheMenu, _struttura.colBlock + dataOreTot - 1]];
            gotoBar.Style = "Top menu GOTO";
            gotoBar.BorderAround2(Weight: Excel.XlBorderWeight.xlMedium, Color: 1);

            //scrivo nome applicazione in alto a sinistra
            Range title = new Range(_struttura.rigaGoto, 2, _struttura.numRigheMenu, _struttura.colBlock - 2);

            int fontSize = 12;
            double rangeSize = _ws.Range[title.ToString()].Width;
            for (; fontSize > 0; fontSize--)
            {
                Graphics grfx = Graphics.FromImage(new Bitmap(1, 1));
                grfx.PageUnit = GraphicsUnit.Point;
                SizeF sizeMax = grfx.MeasureString(Simboli.NomeApplicazione.ToUpper(), new Font("Verdana", fontSize, FontStyle.Bold));
                if (rangeSize > sizeMax.Width)
                    break;
            }

            Style.RangeStyle(_ws.Range[title.ToString()], merge: true, bold: true, fontSize: fontSize, align: Excel.XlHAlign.xlHAlignCenter);
            _ws.Range[title.ToString()].Value = Simboli.NomeApplicazione.ToUpper();

            //calcolo numero elementi per riga
            double numEleRiga = _struttura.numEleMenu / Convert.ToDouble(_struttura.numRigheMenu);

            int j = 0;
            for (int i = 0; i < _struttura.numEleMenu; i++)
            {
                int r = (i / (int)Math.Ceiling(numEleRiga));
                int c = (i % (int)Math.Ceiling(numEleRiga));

                object nome = Struct.tipoVisualizzazione == "O" ? categoriaEntita[i]["SiglaEntita"] : DefinedNames.GetName(categoriaEntita[0]["SiglaEntita"], Date.GetSuffissoData(Workbook.DataAttiva.AddDays(i)));

                Excel.Range rng;
                if (Struct.cell.width.dato < 10)
                {
                    j = c == 0 ? 0 : j + 1;
                    c += j;
                    rng = _ws.Range[_ws.Cells[_struttura.rigaGoto + r, _struttura.colBlock + c + (_struttura.visData0H24 ? 1 : 0)], _ws.Cells[_struttura.rigaGoto + r, _struttura.colBlock + c + 1 + (_struttura.visData0H24 ? 1 : 0)]];
                    rng.Merge();
                }
                else
                {
                    rng = _ws.Cells[_struttura.rigaGoto + r, _struttura.colBlock + c + (_struttura.visData0H24 ? 1 : 0)];
                }

                _definedNames.AddGOTO(nome, Range.R1C1toA1(_struttura.rigaGoto + r, _struttura.colBlock + c + (_struttura.visData0H24 ? 1 : 0)));

                rng.Value = Struct.tipoVisualizzazione == "O" ? categoriaEntita[i]["DesEntitaBreve"] : Workbook.DataAttiva.AddDays(i);
                rng.Style = Struct.tipoVisualizzazione == "O" ? "Barra navigazione con nomi" : "Barra navigazione con date";
            }
            

            //inserisco la data e le ore
            if (Struct.tipoVisualizzazione == "O" || Struct.tipoVisualizzazione == "R")
            {
                int colonnaInizio = _struttura.colBlock;
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    int escludiH24 = (giorno == _dataInizio && _struttura.visData0H24 ? 1 : 0);

                    if (Struct.tipoVisualizzazione == "O")
                    {
                        Range rngData = new Range(_struttura.rigaBlock - 2, colonnaInizio + escludiH24, 1, oreGiorno);

                        Excel.Range rng = _ws.Range[rngData.ToString()];
                        rng.Merge();
                        rng.Style = "Barra della data";
                        rng.Value = giorno.ToString("MM/dd/yyyy");
                        rng.RowHeight = 25;
                    }

                    Range rngOre = new Range(_struttura.rigaBlock - 1, colonnaInizio, 1, oreGiorno + escludiH24);
                    InsertOre(rngOre, giorno == _dataInizio && _struttura.visData0H24);
                    colonnaInizio += oreGiorno + escludiH24;
                });
            }

            ColoraGOTO();
        }
        /// <summary>
        /// Launcher per le azioni di creazione del blocco entità.
        /// </summary>
        /// <param name="entita">La riga contenente le informazioni dell'entità.</param>
        protected void InitBloccoEntita(DataRowView entita)
        {
            SplashScreen.UpdateStatus("Carico struttura " + entita["DesEntita"]);

            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
            DataView grafici = Workbook.Repository[DataBase.TAB.ENTITA_GRAFICO].DefaultView;
            DataView graficiInfo = Workbook.Repository[DataBase.TAB.ENTITA_GRAFICO_INFORMAZIONE].DefaultView;

            if (informazioni.RowFilter != "SiglaEntita = '" + entita["SiglaEntita"] + "' AND IdApplicazione = " + Workbook.IdApplicazione)
            {
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
                graficiInfo.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            }

            if (informazioni.Count == 0)
                return;

            _intervalloOre = Date.GetOreIntervallo(_dataInizio, _dataFine) + (_struttura.visData0H24 ? 1 : 0);// +(_struttura.visParametro ? 1 : 0);

            string funzione = "";
            try
            {
                funzione = "CreaNomiCelle";
                CreaNomiCelle(entita["SiglaEntita"]);
                funzione = "InsertTitoloEntita";
                InsertTitoloEntita(entita["SiglaEntita"], entita["DesEntita"]);
                funzione = "InsertOre";
                InsertOre(entita["SiglaEntita"]);
                funzione = "InsertTitoloVerticale";
                InsertTitoloVerticale(entita["DesEntitaBreve"]);
                funzione = "FormattaBloccoEntita";
                FormattaBloccoEntita();
                funzione = "InsertInformazioniEntita";
                InsertInformazioniEntita();
                funzione = "InsertPersonalizzazioni";
                InsertPersonalizzazioni(entita["SiglaEntita"]);
                funzione = "InsertGrafici";
                InsertGrafici();
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND (ValoreDefault IS NOT NULL OR FormulaInCella = 1) AND IdApplicazione = " + Workbook.IdApplicazione;
                funzione = "InsertFormuleValoriDefault";
                InsertFormuleValoriDefault();
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaTipologiaParametro IS NOT NULL AND IdApplicazione = " + Workbook.IdApplicazione;
                funzione = "InsertParametri";
                InsertParametri();
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                funzione = "FormattazioneCondizionale";
                FormattazioneCondizionale();
            }
            catch (Exception exc)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, Simboli.NomeApplicazione + " Sheet.InitBloccoEntita." + funzione + "[" + entita["SiglaEntita"] + "]");
                throw new LoadStructureException(Simboli.NomeApplicazione + " Sheet.InitBloccoEntita." + funzione + "[" + entita["SiglaEntita"] + "]");
            }
            
            //due righe vuote tra un'entità e la successiva
            _rigaAttiva += 2;
        }
        #region Blocco entità

        /// <summary>
        /// Crea i nomi delle celle in base alla riga. Definisce se sono editabili o meno, se sono parte di selezioni, se devono essere salvate sul DB e se alla modifica deve essere segnalata la modifica stessa o meno. Collega i GOTO generati nel menu con la posizione effettiva delle entità.
        /// </summary>
        /// <param name="siglaEntita">La sigla dell'entità di cui creare i nomi.</param>
        protected virtual void CreaNomiCelle(object siglaEntita)
        {
            //inserisco titoli
            string suffissoData = Date.GetSuffissoData(_dataInizio);
            _definedNames.AddName(_rigaAttiva, Struct.tipoVisualizzazione == "V" ? suffissoData : siglaEntita, "T");

            //sistemo l'indirizzamento dei GOTO
            if (Struct.visLinkEntita)
            {
                int col = _definedNames.GetColFromDate(suffissoData);
                object name = Struct.tipoVisualizzazione == "V" ? DefinedNames.GetName(siglaEntita, suffissoData) : siglaEntita;
                _definedNames.ChangeGOTOAddressTo(name, Range.R1C1toA1(_rigaAttiva, col));
            }

            //aggiungo la riga delle ore
            _rigaAttiva += Struct.tipoVisualizzazione  == "V" ? 2 : 1;

            //aggiungo i grafici
            DataView grafici = Workbook.Repository[DataBase.TAB.ENTITA_GRAFICO].DefaultView;

            int i = 1;
            foreach (DataRowView grafico in grafici)
            {
                _definedNames.AddName(_rigaAttiva++, grafico["SiglaEntita"], "GRAFICO" + i++, Struct.tipoVisualizzazione == "V" ? Date.GetSuffissoData(_dataInizio) : "");
                //i++;
                //_rigaAttiva++;
            }

            //aggiungo informazioni
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            int startCol = _definedNames.GetFirstCol();
            int colOffset = _definedNames.GetColOffset();
            int remove25hour = (Struct.tipoVisualizzazione == "O" ? 0 : 25 - Date.GetOreGiorno(_dataInizio));
            bool isSelection = false;
            string rifSel = "";
            Dictionary<string, int> peers = new Dictionary<string, int>();

            foreach (DataRowView info in informazioni)
            {
                if(Struct.tipoVisualizzazione == "R")
                {
                    CicloGiorni(_dataInizio, _dataInizio.AddDays(_intervalloGiorniMax), (oreGiorno, sufData, giorno) =>
                    {
                        remove25hour = (Struct.tipoVisualizzazione == "O" ? 0 : 25 - Date.GetOreGiorno(giorno));
                        AggiungiNomeInformazione(info, giorno, startCol, colOffset, remove25hour, ref isSelection, ref rifSel, ref peers);
                        _rigaAttiva++;
                    });
                }
                else
                    AggiungiNomeInformazione(info, _dataInizio, startCol, colOffset, remove25hour, ref isSelection, ref rifSel, ref peers);
                

                _rigaAttiva++;
            }
        }

        protected virtual void AggiungiNomeInformazione(DataRowView info, DateTime dataInizio, int startCol, int colOffset, int remove25hour, ref bool isSelection, ref string rifSel, ref Dictionary<string, int> peers)
        {
            object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
            

            _definedNames.AddName(_rigaAttiva, siglaEntitaRif, info["SiglaInformazione"], Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(dataInizio));

            int data0H24 = (info["Data0H24"].Equals("0") && _struttura.visData0H24 ? 1 : 0);

            //selezione - Mantenere in questo ordine: alla prima volta entra nel selezione = 10, poi in isSelection e alla fine chiude la selezione e salta gli altri (se non in presenza di un altro 10)
            if (isSelection && (info["Selezione"].Equals(0) || info["Selezione"].Equals(10)))
            {
                //salvo la selezione
                _definedNames.SetSelection(rifSel, peers);
                //chiudo selezione
                isSelection = false;
                rifSel = "";
                peers = new Dictionary<string, int>();
            }
            if (isSelection)
            {
                Range rng = new Range(_rigaAttiva, startCol - 1);
                peers.Add(rng.ToString(), int.Parse(info["Selezione"].ToString()));
            }
            if (info["Selezione"].Equals(10))
            {
                Range rng = new Range(_rigaAttiva, startCol + data0H24, 1, _definedNames.GetColOffset(_dataFine) - data0H24 - remove25hour);
                isSelection = true;
                rifSel = rng.ToString();
            }
            //fine selezione

            if (info["Editabile"].Equals("1"))
            {
                if (info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
                {
                    //seleziono la cella dell'unità di misura
                    Range rng = new Range(_rigaAttiva, startCol - _visSelezione + 1);
                    _definedNames.SetEditable(_rigaAttiva, rng);
                }
                else
                {
                    Range rng = new Range(_rigaAttiva, startCol + data0H24, 1, _definedNames.GetColOffset(_dataFine) - data0H24 - remove25hour);
                    _definedNames.SetEditable(_rigaAttiva, rng);
                }
            }
            else if (info["Data0H24"].Equals("1"))
            {
                Range rng = new Range(_rigaAttiva, startCol);
                _definedNames.SetEditable(_rigaAttiva, rng);
            }

            if (info["SalvaDB"].Equals("1"))
                _definedNames.SetSaveDB(_rigaAttiva);

            if (info["AnnotaModifica"].Equals("1"))
                _definedNames.SetToNote(_rigaAttiva);

            if (info["SiglaTipologiaInformazione"].Equals("CHECK") && info["Funzione"] != DBNull.Value)
            {
                //cerco parametro n° giorni siglaEntitaRif
                DateTime dataFine = Workbook.DataAttiva.AddDays(Math.Max(
                    (from r in Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].AsEnumerable()
                     where r["IdApplicazione"].Equals(Workbook.IdApplicazione) && r["SiglaEntita"].Equals(siglaEntitaRif) && r["SiglaProprieta"].ToString().EndsWith("GIORNI_STRUTTURA")
                     select int.Parse(r["Valore"].ToString())).FirstOrDefault(), Struct.intervalloGiorni));


                int checkType = int.Parse(Regex.Match(info["Funzione"].ToString(), @"\d+").Value);
                Range rng = new Range(_rigaAttiva, startCol + data0H24, 1, Date.GetOreIntervallo(dataFine) - remove25hour);
                _definedNames.AddCheck(siglaEntitaRif.ToString(), rng.ToString(), checkType);
            }
        }

        /// <summary>
        /// Applica lo stile "Barra titiolo entita" alla riga del titolo dell'entità e scrive la descrizione.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità per la ricerca della riga.</param>
        /// <param name="desEntita">Descrizione da scrivere nella riga.</param>
        protected virtual void InsertTitoloEntita(object siglaEntita, object desEntita)
        {
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                Range rng;
                switch (Struct.tipoVisualizzazione)
                {
                    case "V":
                        rng = _definedNames.Get(suffissoData, "T");
                        break;
                    case "R":
                        rng = _definedNames.Get(siglaEntita, "T");
                        break;
                    default:
                        rng = _definedNames.Get(siglaEntita, "T", suffissoData);
                        break;
                }


                rng.Extend(1, oreGiorno);

                Excel.Range rngTitolo = _ws.Range[rng.ToString()];
                rngTitolo.Merge();
                rngTitolo.Style = "Barra titolo entita";
                rngTitolo.Value = Struct.tipoVisualizzazione == "V" ? giorno.ToString("MM/dd/yyyy") : desEntita.ToString().ToUpperInvariant();
                rngTitolo.RowHeight = 25;
            });
        }
        /// <summary>
        /// In caso di visualizzazione Verticale, formatta e compila la barra delle ore.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità per la ricerca della riga.</param>
        protected virtual void InsertOre(object siglaEntita)
        {
            if (Struct.tipoVisualizzazione == "V")
            {
                Range rng = _definedNames.Get(Date.GetSuffissoData(_dataInizio), "T");
                rng.StartRow++;
                rng.Extend(1, 25);
                InsertOre(rng);
            }
        }
        /// <summary>
        /// Inserisce le ore nel range rng passato per parametro. Se hasData0H24 è true, mette nella prima cella 24.
        /// </summary>
        /// <param name="rng">Range su cui scrivere le ore</param>
        /// <param name="hasData0H24">True se è presente l'ora 24 del giorno precedente.</param>
        private void InsertOre(Range rng, bool hasData0H24 = false)
        {
            Excel.Range rngOre = _ws.Range[rng.ToString()];
            rngOre.Style = "Barra della data";
            rngOre.NumberFormat = "0";
            rngOre.Font.Size = 10;
            rngOre.RowHeight = 20;

            int ora = 1;
            foreach (Range cell in rng.Columns)
            {
                if (hasData0H24) 
                {
                    _ws.Range[cell.ToString()].Value = 24;
                    hasData0H24 = false;
                }
                else
                    _ws.Range[cell.ToString()].Value = ora++;
            }
        }
        /// <summary>
        /// Applica lo stile "Barra titolo verticale" con alcune modifiche e scrive la descrizione dell'entità.
        /// </summary>
        /// <param name="desEntita">Descrizione entità da scrivere.</param>
        protected virtual void InsertTitoloVerticale(object desEntita)
        {
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            Range rngTitolo = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _struttura.colBlock - _visSelezione - 1, Struct.tipoVisualizzazione == "R" ? _intervalloGiorniMax + 1 : informazioni.Count);

            Excel.Range titoloVert = _ws.Range[rngTitolo.ToString()];
            int infoCount = Struct.tipoVisualizzazione == "R" ? _intervalloGiorniMax + 1 : informazioni.Count;

            DataView infoVisible = new DataView(Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE]);
            infoVisible.RowFilter = informazioni.RowFilter + " AND Visibile = '1'";

            string titolo = "";

            Style.RangeStyle(titoloVert, style: "Barra titolo verticale", orientation: infoCount == 1 ? Excel.XlOrientation.xlHorizontal : Excel.XlOrientation.xlVertical, merge: true, fontSize: infoVisible.Count < 5 ? 6 : 9, numberFormat: informazioni.Count > 4 ? "ddd d" : "dd");

            switch (Struct.tipoVisualizzazione)
            {
                case "O":
                    titolo = desEntita.ToString();
                    break;
                case "V":
                    titolo = _dataInizio.ToString();
                    break;
                case "R":
                    titolo = informazioni[0]["DesInformazione"].ToString();
                    break;
            }

            if(titolo.Length > infoVisible.Count)
                titoloVert.Value = "";
            else
                titoloVert.Value = titolo;
        }
        /// <summary>
        /// Formatta l'area dati impostando lo stile di base per le informazioni ("Area dati") e imposta, se ci sono, gli spazi per le informazioni giornaliere.
        /// </summary>
        protected virtual void FormattaBloccoEntita()
        {
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            informazioni.RowFilter += " AND SiglaTipologiaInformazione <> 'GIORNALIERA'";

            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            Range rng = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _definedNames.GetFirstCol() - _visSelezione, Struct.tipoVisualizzazione == "R" ? _intervalloGiorniMax + 1 : informazioni.Count, _definedNames.GetColOffset(_dataFine) + _visSelezione);

            Excel.Range bloccoEntita = _ws.Range[rng.ToString()];
            bloccoEntita.Style = "Area dati";
            bloccoEntita.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            bloccoEntita.Columns[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            bloccoEntita.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            bloccoEntita.Columns[_visSelezione].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
            if (_struttura.visSelezione)
                bloccoEntita.Columns[3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


            int col = _struttura.visData0H24 ? 1 : 0;
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                bloccoEntita.Columns[_visSelezione + col].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                col += oreGiorno;
            });

            informazioni.RowFilter = informazioni.RowFilter.Replace(" AND SiglaTipologiaInformazione <> 'GIORNALIERA'", " AND SiglaTipologiaInformazione = 'GIORNALIERA'");
            if (informazioni.Count > 0)
            {
                rng = new Range(rng.StartRow + rng.RowOffset, rng.StartColumn, informazioni.Count, 2);
                bloccoEntita = _ws.Range[rng.ToString()];
                bloccoEntita.Style = "Area dati";
                bloccoEntita.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
                bloccoEntita.Columns[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                bloccoEntita.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            informazioni.RowFilter = informazioni.RowFilter.Replace(" AND SiglaTipologiaInformazione = 'GIORNALIERA'", "");
        }
        /// <summary>
        /// Inserisce le informazioni e applica la formattazione riga per riga in base alle informazioni sul DB.
        /// </summary>
        protected virtual void InsertInformazioniEntita()
        {
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
            int col = _definedNames.GetFirstCol();
            int colOffset = _definedNames.GetColOffset(_dataFine);
            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            int row = _definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio));

            Excel.Range rngRow = _ws.Range[Range.GetRange(row, col - _visSelezione, informazioni.Count, colOffset + _visSelezione)];
            Excel.Range rngInfo = _ws.Range[Range.GetRange(row, col - _visSelezione, informazioni.Count, 2)];
            Excel.Range rngData = _ws.Range[Range.GetRange(row, col, informazioni.Count, colOffset)];

            if (Struct.tipoVisualizzazione == "V")
            {
                DataView infoNoGiornaliere = new DataView(Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE]);
                infoNoGiornaliere.RowFilter = informazioni.RowFilter + " AND SiglaTipologiaInformazione <> 'GIORNALIERA'";

                Excel.Range rngDataNoGiornaliere = _ws.Range[Range.GetRange(row, col, infoNoGiornaliere.Count, colOffset)];

                int oreGiorno = Date.GetOreGiorno(_dataInizio);
                if(oreGiorno < 24)
                    rngDataNoGiornaliere.Columns[rngDataNoGiornaliere.Columns.Count - 1].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                if(oreGiorno < 25)
                    rngDataNoGiornaliere.Columns[rngDataNoGiornaliere.Columns.Count].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
            }

            int i = 1;
            foreach (DataRowView info in informazioni)
            {
                if (Struct.tipoVisualizzazione == "R")
                {
                    CicloGiorni(_dataInizio, _dataInizio.AddDays(Struct.intervalloGiorni), (oreGiorno, suffissoData, giorno) =>
                    {
                        FormattaInformazione(info, rngInfo.Rows[i], rngRow.Rows[i], rngData.Rows[i], giorno);
                        
                        //disabilito l'ora 24 e 25 dove necessario
                        oreGiorno = Date.GetOreGiorno(giorno);
                        if(oreGiorno < 24)
                            rngData.Rows[i].Columns[rngData.Columns.Count - 1].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                        if(oreGiorno < 25)
                            rngData.Rows[i].Columns[rngData.Columns.Count].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;

                        i++;
                    });
                }
                else 
                {
                    FormattaInformazione(info, rngInfo.Rows[i], rngRow.Rows[i], rngData.Rows[i]);
                    i++;
                }
            }
        }

        protected virtual void FormattaInformazione(DataRowView info, Excel.Range rngInfo, Excel.Range rngRow, Excel.Range rngData, object testoAlternativo = null)
        {
            rngInfo.Value = new object[2] { testoAlternativo != null ? testoAlternativo : info["DesInformazione"], info["DesInformazioneBreve"] };

            int infoBackColor = info["Editabile"].ToString() == "1" ? 15 : 48;

            if (info["Selezione"].Equals(0) && _struttura.visSelezione)
                rngRow.Cells[_visSelezione].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;

            if (info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
            {
                Style.RangeStyle(rngInfo.Cells[1],
                    fontSize: info["FontSize"],
                    foreColor: info["ForeColor"],
                    backColor: (info["Editabile"].ToString() == "1" ? 15 : 48),
                    visible: info["Visibile"].Equals("1"));

                Style.RangeStyle(rngInfo.Cells[2],
                    fontSize: info["FontSize"],
                    foreColor: info["ForeColor"],
                    backColor: info["BackColor"],
                    bold: info["Grassetto"].Equals("1"),
                    numberFormat: info["Formato"],
                    align: Enum.Parse(typeof(Excel.XlHAlign), info["Align"].ToString()));
            }
            else if (info["SiglaTipologiaInformazione"].Equals("TITOLO2"))
            {
                //caselle delle informazioni + selezione + DATA0H24
                Range rng = new Range(rngInfo.Address).ExtendOf(colOffset: _struttura.visData0H24 ? 1 : 0).ExtendOf(colOffset: _struttura.visSelezione ? 1 : 0);

                Style.RangeStyle(_ws.Range[rng.ToString()],
                    fontSize: info["FontSize"],
                    foreColor: info["ForeColor"],
                    backColor: info["BackColor"],
                    merge: true,
                    bold: true,
                    borders: "[Top:medium, Right:medium]",
                    visible: info["Visibile"].Equals("1"));

                int col = _struttura.visData0H24 ? 1 : 0;
                //giorni normali
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    rng = new Range(rngData.Address);
                    Style.RangeStyle(_ws.Range[rng.Columns[col, col + oreGiorno - 1].ToString()],
                        fontSize: info["FontSize"],
                        foreColor: info["ForeColor"],
                        backColor: info["BackColor"],
                        align: Excel.XlHAlign.xlHAlignCenter,
                        merge: true,
                        bold: true,
                        borders: "[Top:medium, Right:medium]");

                    _ws.Range[rng.Columns[col, col + oreGiorno - 1].ToString()].Value = info["DesInformazione"];

                    col += oreGiorno;
                });
            }
            else
            {
                if (info["InizioGruppo"].Equals("1"))
                    rngRow.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;

                Style.RangeStyle(rngInfo,
                    fontSize: info["FontSize"],
                    foreColor: info["ForeColor"],
                    backColor: infoBackColor,
                    visible: info["Visibile"].Equals("1"),
                    numberFormat: Struct.tipoVisualizzazione == "R" ? "dd/MM/yyyy" : "general",
                    borders: "[Right:medium]");

                Style.RangeStyle(rngData,
                    fontSize: info["FontSize"],
                    foreColor: info["ForeColor"],
                    backColor: info["BackColor"],
                    bold: info["Grassetto"].Equals("1"),
                    numberFormat: info["Formato"],
                    align: Enum.Parse(typeof(Excel.XlHAlign), info["Align"].ToString()));

                if (info["Data0H24"].Equals("0") && _struttura.visData0H24 && !info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
                    rngData.Cells[1].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;

                if (_struttura.visData0H24)
                    rngData.Cells[1].Font.Size -= 1;

                //Celle merged
                if (info["CriterioUnione"] != DBNull.Value)
                {
                    string[] uCriteria = info["CriterioUnione"].ToString().Split(';');
                    foreach (string uCrit in uCriteria)
                    {
                        var interval = uCrit
                            .Split('-')
                            .OfType<string>()
                            .Select(s => int.Parse(Regex.Match(s, @"\d+").Value))
                            .ToArray();

                        if (interval[1] == 25 && Date.GetOreGiorno(_dataInizio) < 25)
                            interval[1] = Date.GetOreGiorno(_dataInizio);

                        Range union = new Range(rngData.Columns[interval[0]].Address() + ":" + rngData.Columns[interval[1]].Address());

                        _ws.Range[union.ToString()].Merge();
                    }
                    rngData.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                }

                //Validazione celle se UNIT_COMM
                if (info["SiglaInformazione"].Equals("UNIT_COMM"))
                {
                    var unit_comm = Workbook.Repository[DataBase.TAB.ENTITA_COMMITMENT]
                        .AsEnumerable()
                        .Where(r => r["SiglaEntita"].Equals(info["SiglaEntitaRif"]) || r["SiglaEntita"].Equals(info["SiglaEntita"]))
                        .ToList();

                    if (unit_comm.Count > 0)
                    {
                        Range rng = new Range(rngData.Address);
                        rng.StartColumn += _struttura.visData0H24 ? 1 : 0;

                        string formula = "=O(";
                        string valoriAmmessi = "";
                        foreach (DataRow r in unit_comm)
                        {
                            formula += rng.Cells[0, 0].ToString() + "=\"" + r["SiglaCommitment"] + "\";";
                            valoriAmmessi += r["SiglaCommitment"].ToString() + ", ";
                        }
                        formula = formula.Substring(0, formula.Length - 1) + ")";
                        valoriAmmessi = valoriAmmessi.Substring(0, valoriAmmessi.Length - 2);

                        Excel.Validation v = _ws.Range[rng.ToString()].Validation;
                        v.Delete();
                        v.Add(Type: Excel.XlDVType.xlValidateCustom,
                            AlertStyle: Excel.XlDVAlertStyle.xlValidAlertStop,
                            Formula1: formula);
                        v.IgnoreBlank = false;
                        v.InputTitle = "Unit Commitment";
                        v.InputMessage = "Digitare un valore tra i seguenti: " + valoriAmmessi;
                        v.ErrorTitle = "Valore non ammesso";
                        v.ErrorMessage = "Il valore digitato non è tra quelli ammessi per lo Unit Commitment di questa UP. Sceglierne uno tra i seguenti: " + valoriAmmessi;
                        v.ShowError = true;
                        v.ShowInput = true;

                        Marshal.ReleaseComObject(v);
                        v = null;
                    }
                }
                else if (info["SiglaInformazione"].Equals("RISPETTO_PROG_PREC"))
                {
                    Range rng = new Range(rngData.Address);
                    rng.StartColumn += _struttura.visData0H24 ? 1 : 0;


                    string cell = rng.Cells[0, 0].ToString();
                    string formula = "=O(" + cell + "=\"SI\";" + cell + "=\"NO\")";
                    string valoriAmmessi = "SI, NO";

                    Excel.Validation v = _ws.Range[rng.ToString()].Validation;
                    v.Delete();
                    v.Add(Type: Excel.XlDVType.xlValidateCustom,
                        AlertStyle: Excel.XlDVAlertStyle.xlValidAlertStop,
                        Formula1: formula);
                    v.IgnoreBlank = false;
                    v.InputTitle = "Rispetto a progr. prec.";
                    v.InputMessage = "Digitare un valore tra i seguenti: " + valoriAmmessi;
                    v.ErrorTitle = "Valore non ammesso";
                    v.ErrorMessage = "Il valore digitato non è tra quelli ammessi per 'Rispetto a prog. prec.' di questa UP. Sceglierne uno tra i seguenti: " + valoriAmmessi;
                    v.ShowError = true;
                    v.ShowInput = true;

                    Marshal.ReleaseComObject(v);
                    v = null;
                }

                if (info["DesInformazione"].Equals("ACQ/VEN"))
                {
                    Range rng = new Range(rngData.Address);
                    rng.StartColumn += _struttura.visData0H24 ? 1 : 0;
                        

                    string cell = rng.Cells[0, 0].ToString();
                    string formula = "=O(" + cell + "=\"ACQ\";" + cell + "=\"VEN\")";
                    string valoriAmmessi = "ACQ, VEN";

                    Excel.Validation v = _ws.Range[rng.ToString()].Validation;
                    v.Delete();
                    v.Add(Type: Excel.XlDVType.xlValidateCustom,
                        AlertStyle: Excel.XlDVAlertStyle.xlValidAlertStop,
                        Formula1: formula);
                    v.IgnoreBlank = false;
                    v.InputTitle = "ACQ/VEN";
                    v.InputMessage = "Digitare un valore tra i seguenti: " + valoriAmmessi;
                    v.ErrorTitle = "Valore non ammesso";
                    v.ErrorMessage = "Il valore digitato non è tra quelli ammessi per VEN/ACQ di questa UP. Sceglierne uno tra i seguenti: " + valoriAmmessi;
                    v.ShowError = true;
                    v.ShowInput = true;

                    Marshal.ReleaseComObject(v);
                    v = null;
                }
            }
            
        }
        /// <summary>
        /// Inserisce i valori di default e le formule.
        /// </summary>
        protected virtual void InsertFormuleValoriDefault()
        {
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
            int colOffset = Date.GetOreIntervallo(_dataFine);

            foreach (DataRowView info in informazioni)
            {
                if (Struct.tipoVisualizzazione == "R")
                {
                    CicloGiorni(_dataInizio, _dataInizio.AddDays(_intervalloGiorniMax), (oreGiorno, suffissoData, giorno) =>
                    {
                        AddInformazioneValue(info, giorno, colOffset);
                    });
                }
                else
                {
                    CicloGiorni(_dataInizio, _dataFine, (oreGiorno, suffissoData, giorno) =>
                    {
                        AddInformazioneValue(info, giorno, oreGiorno);
                    });
                    //AddInformazioneValue(info, _dataInizio, colOffset);
                    if (info["SiglaTipologiaInformazione"].Equals("OTTIMO"))
                    {
                        AddOptFunction(info, _dataInizio, colOffset);
                    }
                }
            }
        }

        protected virtual void AddOptFunction(DataRowView info, DateTime giorno, int colOffset)
        {
            object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

            Range rng = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], Date.GetSuffissoData(giorno)), _definedNames.GetColData1H1(), 1, colOffset);
            
            Range data0H24 = _struttura.visData0H24 && info["Data0H24"].Equals("1") ? new Range(rng.StartRow, _definedNames.GetFirstCol()) : null;

            _ws.Range[data0H24.ToString()].Formula = "=SUM(" + rng.Columns[0, rng.Columns.Count - 1] + ")";
        }

        protected virtual void AddInformazioneValue(DataRowView info, DateTime giorno, int colOffset)
        {
            object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

            Range rng = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], Date.GetSuffissoData(giorno)), _definedNames.GetColFromDate(Date.GetSuffissoData(giorno)));
            Range data0H24 = _struttura.visData0H24 && info["Data0H24"].Equals("1") ? new Range(rng.StartRow, _definedNames.GetFirstCol()) : null;

            if (info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
                rng.StartColumn -= _visSelezione - 1;
            else
                rng.Extend(colOffset: colOffset);

            Excel.Range rngData = _ws.Range[rng.ToString()];

            if (info["ValoreDefault"] != DBNull.Value)
            {
                rngData.Value = info["ValoreDefault"];
            }
            else if (info["FormulaInCella"].Equals("1"))
            {
                int deltaNeg;
                int deltaPos;
                string formula = "=" + PreparaFormula(info, Date.GetSuffissoData(giorno.AddDays(-1)), Date.GetSuffissoData(giorno), 24, out deltaNeg, out deltaPos);

                //if (info["SiglaTipologiaInformazione"].Equals("OTTIMO"))
                //    _ws.Range[data0H24.ToString()].Formula = "=SUM(" + rng.Columns[0, rng.Columns.Count - 1] + ")";

                _ws.Range[rng.Columns[deltaNeg, rng.Columns.Count - 1 - deltaPos].ToString()].Formula = formula;
            }

            if (data0H24 != null && info["ValoreData0H24"] != DBNull.Value)
                _ws.Range[data0H24.ToString()].Value = info["ValoreData0H24"];
        }
        /// <summary>
        /// Inserisce i parametri.
        /// </summary>
        protected virtual void InsertParametri()
        {
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
            //DataView parametriD = Workbook.Repository[DataBase.TAB.ENTITA_PARAMETRO_D].DefaultView;
            //DataView parametriH = Workbook.Repository[DataBase.TAB.ENTITA_PARAMETRO_H].DefaultView;
            DataView parametri = Workbook.Repository[DataBase.TAB.ENTITA_PARAMETRO].DefaultView;

            foreach (DataRowView info in informazioni)
            {
                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                DateTime dataFine = Struct.tipoVisualizzazione == "R" ? _dataInizio.AddDays(_intervalloGiorniMax) : _dataFine;

                CicloGiorni(_dataInizio, dataFine, (oreGiorno, suffissoData, giorno) =>
                {

                    Range rngData = _definedNames.Get(siglaEntita, info["SiglaInformazione"], suffissoData).Extend(colOffset: oreGiorno);

                    parametri.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaParametro = '" + info["SiglaTipologiaParametro"] + "' AND DataIV <= '" + giorno.ToString("yyyyMMdd") + "01' AND DataFV >= '" + giorno.ToString("yyyyMMdd") + "25' AND IdApplicazione = " + Workbook.IdApplicazione;

                    if (parametri.Count == 1) 
                    {
                        _ws.Range[rngData.ToString()].Formula = parametri[0]["Valore"];
                    }
                    else
                    {
                        for (int i = 1; i <= oreGiorno; i++)
                        {
                            parametri.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaParametro = '" + info["SiglaTipologiaParametro"] + "' AND DataIV <= '" + giorno.ToString("yyyyMMdd") + i.ToString("00") + "' AND DataFV >= '" + giorno.ToString("yyyyMMdd") + i.ToString("00") + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                            
                            if (parametri.Count == 1)
                                _ws.Range[rngData.Columns[i - 1].ToString()].Value = parametri[0]["Valore"];
                        }

                        //object[] values = parametri.ToTable(false, "Valore").AsEnumerable().Select(r => r["Valore"]).ToArray();

                        //if (values.Length > 0)
                        //    _ws.Range[rngData.ToString()].Value = values;
                    }
                });
            }
        }
        /// <summary>
        /// Crea la formattazione condizionale.
        /// </summary>
        protected virtual void FormattazioneCondizionale()
        {
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
            DataView formattazione = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE_FORMATTAZIONE].DefaultView;
            int colOffset = _definedNames.GetColOffset(_dataFine);
            foreach (DataRowView info in informazioni)
            {
                if (Struct.tipoVisualizzazione == "R")
                {
                    CicloGiorni(_dataInizio, _dataInizio.AddDays(_intervalloGiorniMax), (oreGiorno, suffissoData, giorno) =>
                    {
                        AddFormattazioneInformazione(info, giorno, formattazione, colOffset);
                    });
                }
                else
                    AddFormattazioneInformazione(info, _dataInizio, formattazione, colOffset);
            }
        }

        protected virtual void AddFormattazioneInformazione(DataRowView info, DateTime giorno, DataView formattazione, int colOffset)
        {
            object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

            int offsetAdjust = (_struttura.visData0H24 && info["Data0H24"].Equals("0") ? 1 : 0);
            Range rng = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], Date.GetSuffissoData(giorno)), _definedNames.GetFirstCol() + offsetAdjust, 1, colOffset - offsetAdjust);

            Excel.Range rngData = _ws.Range[rng.ToString()];

            formattazione.RowFilter = (info["SiglaEntitaRif"] is DBNull ? "SiglaEntita" : "SiglaEntitaRif") + " = '" + siglaEntita + "' AND SiglaInformazione = '" + info["SiglaInformazione"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            foreach (DataRowView format in formattazione)
            {

                string[] valore = format["Valore"].ToString().Replace("\"", "").Split('|');
                if (format["NomeCella"] != DBNull.Value)
                {
                    int refRow = _definedNames.GetRowByNameSuffissoData(siglaEntita, format["NomeCella"], Date.GetSuffissoData(giorno));
                    string address = Range.GetRange(refRow, rng.StartColumn);
                    string formula = "";
                    switch ((int)format["Operatore"])
                    {
                        case 1:
                            formula = "=E(" + address + ">=" + valore[0] + ";" + address + "<=" + valore[1] + ")";
                            break;
                        case 3:
                            formula = "=" + address + "=" + valore[0];
                            break;
                        case 5:
                            formula = "=" + address + ">" + valore[0];
                            break;
                        case 6:
                            formula = "=" + address + "<" + valore[0];
                            break;
                    }

                    Excel.FormatCondition cond = rngData.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: formula);

                    cond.Font.Color = format["ForeColor"];
                    cond.Font.Bold = format["Grassetto"].Equals("1");
                    if ((int)format["BackColor"] == 0)
                        cond.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    else
                        cond.Interior.Color = format["BackColor"];
                    cond.Interior.Pattern = format["Pattern"];
                }
                else
                {
                    string formula1;
                    string formula2 = "";
                    if ((int)format["Operatore"] == 1)
                    {
                        formula1 = valore[0];
                        formula2 = valore[1];
                    }
                    else
                    {
                        formula1 = valore[0];
                    }

                    Excel.FormatCondition cond = rngData.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, format["Operatore"], formula1, formula2);

                    cond.Font.Color = format["ForeColor"];
                    cond.Font.Bold = format["Grassetto"].Equals("1");
                    if ((int)format["BackColor"] == 0)
                        cond.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    else
                        cond.Interior.Color = format["BackColor"];

                    cond.Interior.Pattern = format["Pattern"];
                }
            }
        }

        /// <summary>
        /// Inserisce i grafici creando anche le serie.
        /// </summary>
        protected override void InsertGrafici()
        {
            DataView grafici = Workbook.Repository[DataBase.TAB.ENTITA_GRAFICO].DefaultView;
            DataView graficiInfo = Workbook.Repository[DataBase.TAB.ENTITA_GRAFICO_INFORMAZIONE].DefaultView;

            int i = 1;
            int col = _definedNames.GetColData1H1();
            int colOffset = _definedNames.GetColOffset(_dataFine) - (_struttura.visData0H24 ? 1 : 0);
            foreach (DataRowView grafico in grafici)
            {
                SplashScreen.UpdateStatus("Genero grafici");
                string name = DefinedNames.GetName(grafico["SiglaEntita"], "GRAFICO" + i++, Struct.tipoVisualizzazione == "V" ? Date.GetSuffissoData(_dataInizio) : "");

                Range rngGrafico = new Range(_definedNames.GetRowByName(name), col, 1, colOffset);
                Excel.Range xlRngGrafico = _ws.Range[rngGrafico.ToString()];
                xlRngGrafico.Merge();
                xlRngGrafico.Style = "Area grafici";
                xlRngGrafico.RowHeight = 200;
                Excel.Chart chart = _ws.ChartObjects().Add(Left: xlRngGrafico.Left, Top: xlRngGrafico.Top + 1, Width: xlRngGrafico.Width, Height: xlRngGrafico.Height - 2).Chart;

                chart.Parent.Name = name;

                chart.Axes(Excel.XlAxisType.xlCategory).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
                chart.Axes(Excel.XlAxisType.xlValue).HasMajorGridlines = false;
                chart.Axes(Excel.XlAxisType.xlValue).HasMinorGridlines = false;
                chart.Axes(Excel.XlAxisType.xlValue).MinorTickMark = Excel.XlTickMark.xlTickMarkOutside;
                chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Name = "Verdana";
                chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Size = 11;
                chart.Axes(Excel.XlAxisType.xlValue).TickLabels.NumberFormat = "general";

                chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
                chart.HasDataTable = false;
                chart.DisplayBlanksAs = Excel.XlDisplayBlanksAs.xlNotPlotted;
                chart.ChartGroups(1).GapWidth = 0;
                chart.ChartGroups(1).Overlap = 100;
                chart.ChartArea.Border.ColorIndex = 1;
                chart.ChartArea.Border.Weight = 3;
                chart.ChartArea.Border.LineStyle = 0;
                chart.PlotVisibleOnly = false;

                chart.PlotArea.Top = chart.ChartArea.Height;

                chart.PlotArea.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                string rowFilter = graficiInfo.RowFilter;
                graficiInfo.RowFilter = rowFilter + " AND SiglaGrafico = '" + grafico["SiglaGrafico"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;

                bool hasSecondaryAxes = false;
                foreach (DataRowView info in graficiInfo)
                {
                    if (Struct.tipoVisualizzazione == "R")
                    {
                        CicloGiorni(_dataInizio, _dataInizio.AddDays(Struct.intervalloGiorni), (oreGiorno, suffissoData, giorno) =>
                        {
                            Range rngDati = new Range(_definedNames.GetRowByNameSuffissoData(grafico["SiglaEntita"], info["SiglaInformazione"], Date.GetSuffissoData(giorno)), col, 1, Date.GetOreGiorno(giorno));
                            Excel.Series serie = chart.SeriesCollection().NewSeries();
                            serie.Name = giorno.ToString("dd/MM/yyyy");
                            serie.Values = _ws.Range[rngDati.ToString()];
                            serie.ChartType = (Excel.XlChartType)info["ChartType"];
                        });
                    }
                    else
                    {
                        Range rngDati = new Range(_definedNames.GetRowByNameSuffissoData(grafico["SiglaEntita"], info["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), col, 1, colOffset);
                        Excel.Series serie = chart.SeriesCollection().NewSeries();
                        serie.Name = info["DesInformazione"].ToString();
                        serie.Values = _ws.Range[rngDati.ToString()];
                        serie.ChartType = (Excel.XlChartType)info["ChartType"];
                        serie.Interior.ColorIndex = info["InteriorColor"];
                        serie.Border.ColorIndex = info["BorderColor"];
                        serie.Border.Weight = info["BorderWeight"];
                        serie.Border.LineStyle = info["BorderLineStyle"];
                        serie.AxisGroup = (Excel.XlAxisGroup)info["AxisGroup"];
                        if ((Excel.XlAxisGroup)info["AxisGroup"] == Excel.XlAxisGroup.xlSecondary)
                            hasSecondaryAxes = true;
                    }

                    if (info["ScalaMin"] != DBNull.Value)
                        chart.Axes(Excel.XlAxisType.xlValue, (Excel.XlAxisGroup)info["AxisGroup"]).MinimumScale = info["ScalaMin"];
                    if (info["ScalaMax"] != DBNull.Value)
                        chart.Axes(Excel.XlAxisType.xlValue, (Excel.XlAxisGroup)info["AxisGroup"]).MaximumScale = info["ScalaMax"];
                    if (info["Intervallo"] != DBNull.Value)
                        chart.Axes(Excel.XlAxisType.xlValue, (Excel.XlAxisGroup)info["AxisGroup"]).MajorUnit = info["Intervallo"];
                }


                //asse secondario
                if (hasSecondaryAxes)
                {
                    chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).HasMajorGridlines = false;
                    chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).HasMinorGridlines = false;
                    chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).MinorTickMark = Excel.XlTickMark.xlTickMarkOutside;
                    chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).TickLabels.Font.Name = "Verdana";
                    chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).TickLabels.Font.Size = 11;
                    chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).TickLabels.NumberFormat = "general";
                }

                graficiInfo.RowFilter = rowFilter;
            }
        }
        /// <summary>
        /// Aggiorna tutti i grafici del foglio.
        /// </summary>
        public override void AggiornaGrafici()
        {
            if (_ws.ChartObjects().Count > 0)
            {
                ((Excel._Worksheet)_ws).Calculate();
                Excel.ChartObjects charts = _ws.ChartObjects();
                foreach (Excel.ChartObject chart in charts)
                {
                    int col;
                    if (chart.Name.Contains("DATA"))
                    {
                        col = _definedNames.GetColFromDate(chart.Name.Split(Simboli.UNION[0]).Last());
                    }
                    else
                    {
                        col = _definedNames.GetColFromDate();
                    }
                    int row = _definedNames.GetRowByName(chart.Name);
                    Excel.Range rng = _ws.Range[Range.GetRange(row, col)];
                    AggiornaGrafici(chart.Chart, rng.MergeArea);
                    //chart.Chart.Refresh();
                }
            }
        }
        /// <summary>
        /// Allinea il grafico al range in modo da far combaciare la barra delle ordinate con la prima colonna dell'area dati. Per far questo calcola la dimensione in punti dei label di ordinata e sposta di conseguenza l'area del grafico.
        /// </summary>
        /// <param name="chart">Microsoft.Office.Interop.Excel.Chart da aggiornare.</param>
        /// <param name="rigaGrafico">Microsoft.Office.Interop.Excel.Range a cui il grafico appartiene.</param>
        private void AggiornaGrafici(Excel.Chart chart, Excel.Range rigaGrafico)
        {
            SplashScreen.UpdateStatus("Aggiorno grafici " + chart.Name);

            chart.Refresh();
            //resize dell'area del grafico per adattarla alle ore
            using (Graphics grfx = Graphics.FromImage(new Bitmap(1, 1)))
            {
                grfx.PageUnit = GraphicsUnit.Point;
                grfx.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;                
                float sizeMax = float.MinValue;
                SizeF tmpSize;
                double val = chart.Axes(Excel.XlAxisType.xlValue).MinimumScale;

                //controllo anche il fondo scala: se cambia l'ordine di grandezza excel lascia lo spazio nel label come se ci fosse!!
                while (val <= chart.Axes(Excel.XlAxisType.xlValue).MaximumScale)
                {
                    tmpSize = grfx.MeasureString(val.ToString(), new Font(chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Name, (float)chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Size));
                    sizeMax = Math.Max(sizeMax, tmpSize.Width);

                    val += chart.Axes(Excel.XlAxisType.xlValue).MajorUnit;
                }

                //MANTENERE ORDINE DI QUESTE ISTRUZIONI
                chart.ChartArea.Left = rigaGrafico.Left - Math.Ceiling(sizeMax) - 7;      //sposto a destra il grafico
                chart.ChartArea.Width = rigaGrafico.Width + Math.Ceiling(sizeMax) + 4;    //aumento la larghezza del grafico
                Excel.PlotArea plotArea = chart.PlotArea;
                try
                {
                    plotArea.InsideLeft = 0d;                                               //allineo il grafico al bordo sinistro dell'area esterna al grafico
                }
                catch { }
                plotArea.Width = chart.ChartArea.Width + 3;                                 //aumento la larghezza dell'area esterna al grafico
                Marshal.ReleaseComObject(plotArea);
                plotArea = null;
                
                //se c'è un asse secondario ed aggiorno la dimensione del grafico di conseguenza
                try
                {
                    sizeMax = float.MinValue;
                    val = 0;
                    while (val < chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).MaximumScale)
                    {
                        if ((val += chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).MajorUnit) > chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).MaximumScale)
                            val = chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).MaximumScale;

                        tmpSize = grfx.MeasureString(val.ToString(), new Font(chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).TickLabels.Font.Name, (float)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).TickLabels.Font.Size));
                        sizeMax = Math.Max(sizeMax, tmpSize.Width);
                    }

                    chart.ChartArea.Width = chart.ChartArea.Width + Math.Ceiling(sizeMax) + 7;    //aumento la larghezza del grafico

                } catch {}

                //se visualizzazione R blocco il grafico alla 24
                if (Struct.tipoVisualizzazione == "R")
                {
                    bool start = TimeZone.CurrentTimeZone.IsDaylightSavingTime(Workbook.DataAttiva);
                    bool end = TimeZone.CurrentTimeZone.IsDaylightSavingTime(Workbook.DataAttiva.AddDays(Struct.intervalloGiorni));
                    
                    if(!start || end)
                        chart.ChartArea.Width -= _ws.Range[Range.GetRange(1, _definedNames.GetColFromDate(Date.SuffissoDATA1, Date.GetSuffissoOra(25)))].Width;
                }
            }
            chart.Refresh();
        }

        #endregion

        /// <summary>
        /// Launcher per caricare le informazioni e i commenti dal DB.
        /// </summary>
        public override void CaricaInformazioni()
        {
            try
            {
                //inserire qui la gestione del cambio mercato per MI
                if (DataBase.OpenConnection())
                {
                    DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                    categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND IdApplicazione = " + Workbook.IdApplicazione;

                    _dataInizio = Workbook.DataAttiva;

                    DateTime dataFineMax = _dataInizio.AddDays(_intervalloGiorniMax);

                    DataView datiApplicazione = Workbook.Repository[DataBase.TAB.DATI_APPLICAZIONE].DefaultView;
                    DataView insertManuali = Workbook.Repository[DataBase.TAB.DATI_APPLICAZIONE_COMMENTO].DefaultView;

                    if (Struct.tipoVisualizzazione == "O")
                    {
                        foreach (DataRowView entita in categoriaEntita)
                        {
                            object siglaEntita = entita["Gerarchia"] is DBNull ? entita["SiglaEntita"] : entita["Gerarchia"];
                            SplashScreen.UpdateStatus("Scrivo informazioni " + entita["DesEntita"]);
                            datiApplicazione.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(Data, System.Int32) <= " + _dataFineUP[siglaEntita].ToString("yyyyMMdd");
                            CaricaInformazioniEntita(datiApplicazione);
                            insertManuali.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(SUBSTRING(Data, 1, 8), System.Int32) <= " + _dataFineUP[siglaEntita].ToString("yyyyMMdd");
                            CaricaCommentiEntita(insertManuali);
                        }
                    }
                    else
                    {                        
                        CaricaInformazioniEntita(datiApplicazione);
                        CaricaCommentiEntita(insertManuali);
                    }
                    //TODO se dati giornalieri riabilitare
                    //SplashScreen.UpdateStatus("Carico dati giornalieri");
                    ////carico dati giornalieri
                    //DataView datiApplicazioneD = Workbook.Repository[DataBase.TAB.DATI_APPLICAZIONE_D].DefaultView;

                    //foreach (DataRowView dato in datiApplicazioneD)
                    //{
                    //    if (_definedNames.IsDefined(dato["SiglaEntita"]))
                    //    {
                    //        Range rng = new Range(_definedNames.GetRowByNameSuffissoData(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(dato["Data"].ToString())), _definedNames.GetFirstCol() - 1);

                    //        _ws.Range[rng.ToString()].Value = dato["Valore"];
                    //    }
                    //}
                }
            }
            catch (Exception e)
            {
                Workbook.InsertLog(PSO.Core.DataBase.TipologiaLOG.LogErrore, "CaricaInformazioni [all = 1]: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Carica le informazioni.
        /// </summary>
        /// <param name="datiApplicazione">Tabella contenente tutte le informazioni da scrivere.</param>
        protected virtual void CaricaInformazioniEntita(DataView datiApplicazione)
        {
            foreach (DataRowView dato in datiApplicazione)
            {
                DateTime giorno = DateTime.ParseExact(dato["Data"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                
                if (_dataFineUP.ContainsKey(dato["SiglaEntita"]) && giorno <= _dataFineUP[dato["SiglaEntita"]])
                {
                    //sono nel caso DATA0H24
                    if (giorno < Workbook.DataAttiva)
                    {
                        Range rng = _definedNames.Get(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(Workbook.DataAttiva.AddDays(-1)), Date.GetSuffissoOra(24));
                        _ws.Range[rng.ToString()].Value = dato["H24"];
                    }
                    else
                    {
                        Range rng = _definedNames.Get(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(giorno)).Extend(colOffset: Date.GetOreGiorno(giorno));
                        List<object> o = new List<object>(dato.Row.ItemArray);
                        _ws.Range[rng.ToString()].Value = o.ToArray();

                        if (giorno == Workbook.DataAttiva && Regex.IsMatch(dato["SiglaInformazione"].ToString(), @"RIF\d+"))
                        {
                            Selection s = _definedNames.GetSelectionByRif(rng);
                            s.ClearSelections(_ws);
                            s.Select(_ws, int.Parse(o[0].ToString().Split('.')[0]));
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Carica i commenti.
        /// </summary>
        /// <param name="insertManuali">Tabella che contiene tutte le informazioni che necessitano del commento</param>
        protected virtual void CaricaCommentiEntita(DataView insertManuali)
        {
            foreach (DataRowView commento in insertManuali)
            {
                DateTime giorno = DateTime.ParseExact(commento["Data"].ToString().Substring(0,8), "yyyyMMdd", CultureInfo.InvariantCulture);

                if (_dataFineUP.ContainsKey(commento["SiglaEntita"]) && giorno <= _dataFineUP[commento["SiglaEntita"]])
                {
                    SplashScreen.UpdateStatus("Scrivo commenti " + commento["SiglaEntita"]);
                    Range rng = _definedNames.Get(commento["SiglaEntita"], commento["SiglaInformazione"], Date.GetSuffissoData(giorno), Date.GetSuffissoOra(commento["Data"].ToString()));
                    _ws.Range[rng.ToString()].ClearComments();
                    _ws.Range[rng.ToString()].AddComment("Valore inserito manualmente");
                }
            }
        }        

        /// <summary>
        /// Prepara la formula per essere scritta in cella. Sostituisce i parametri con i riferimenti delle celle corrispondenti. Calcola anche un possibile offset se la formula va a controllare valori delle ore precedenti o successive.
        /// </summary>
        /// <param name="info">La riga dell'informazione.</param>
        /// <param name="suffissoDataPrec">Il suffisso della data antecedente a quella in cui si sta lavorando (solitamente DATA0).</param>
        /// <param name="suffissoData">Il suffisso della data in cui si sta lavorando (solitamente DATA1).</param>
        /// <param name="oreDataPrec">Numero di ore della data precedente.</param>
        /// <param name="deltaNeg">Parametro di output che indica l'offset dall'inizio del giorno.</param>
        /// <param name="deltaPos">Parametro di output che indica l'offset dalla fine del giorno.</param>
        /// <returns></returns>
        protected string PreparaFormula(DataRowView info, string suffissoDataPrec, string suffissoData, int oreDataPrec, out int deltaNeg, out int deltaPos)
        {
            if (info["Formula"] != DBNull.Value || info["Funzione"] != DBNull.Value)
            {
                string formula = info["Formula"] is DBNull ? info["Funzione"].ToString() : info["Formula"].ToString();

                string[] parametri = info["FormulaParametro"].ToString().Split(',');

                int tmpdeltaNeg = 0;
                int tmpdeltaPos = 0;

                foreach (string par in parametri)
                {
                    if (Regex.IsMatch(par, @"\[[-+]?\d+\]"))
                    {
                        int deltaOre = int.Parse(par.Split('[')[1].Replace("]", ""));
                        if (deltaOre > 0)
                            tmpdeltaPos = Math.Max(tmpdeltaPos, deltaOre);
                        else
                            tmpdeltaNeg = Math.Min(tmpdeltaNeg, deltaOre);
                    }
                }

                deltaNeg = tmpdeltaNeg != 0 ? Math.Abs(tmpdeltaNeg + 1) : 0;
                deltaPos = tmpdeltaPos;

                formula = Regex.Replace(formula, @"%P\d+(E\d+)?(\$(\d+|I|F))?%", delegate(Match m)
                {
                    string[] parametroEntitaLock = m.Value.Split('$');

                    int oraLock = -1;
                    if (parametroEntitaLock.Length > 1)
                    {
                        if (parametroEntitaLock[1] == "f%" || parametroEntitaLock[1] == "F%")
                            oraLock = Date.GetOreGiorno(suffissoData);
                        else if (parametroEntitaLock[1] == "i%" || parametroEntitaLock[1] == "I%")
                            oraLock = 1;
                        else
                            int.TryParse(Regex.Match(parametroEntitaLock[1], @"\d+").Value, out oraLock);
                    }

                    string[] parametroEntita = parametroEntitaLock[0].Split('E');
                    
                    int n = int.Parse(Regex.Match(parametroEntita[0], @"\d+").Value);

                    object siglaEntita = "";
                    string siglaInformazione = "";
                    string suffData = "";
                    string suffOra = "";
                    if (parametroEntita.Length > 1)
                    {
                        int eRif = int.Parse(Regex.Match(parametroEntita[1], @"\d+").Value);
                        DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                        categoriaEntita.RowFilter = "Gerarchia = '" + info["SiglaEntita"] + "' AND Riferimento = " + eRif + " AND IdApplicazione = " + Workbook.IdApplicazione;
                        if (categoriaEntita.Count == 0)
                            categoriaEntita.RowFilter = "Riferimento = " + eRif + " AND IdApplicazione = " + Workbook.IdApplicazione;
                        siglaEntita = categoriaEntita[0]["SiglaEntita"];
                    }
                    else
                        siglaEntita = info["SiglaEntita"];
                    
                    siglaInformazione = parametri[n - 1];
                    
                    if (Regex.IsMatch(siglaInformazione, @"\[[-+]?\d+\]"))
                    {
                        int deltaOre = int.Parse(siglaInformazione.Split('[')[1].Replace("]", ""));

                        if (suffissoData == "DATA1")
                        {//traslo in avanti la formula di |deltaNeg| - |deltaOre|
                            int ora = Math.Abs(tmpdeltaNeg) + deltaOre;// +(info["Data0H24"].Equals("1") ? 0 : 1);
                            suffData = ora == 0 ? "DATA0" : "DATA1";
                            suffOra = ora == 0 ? "H24" : "H" + ora;
                        }
                        else
                        {
                            int ora = (deltaOre < 0 ? oreDataPrec + deltaOre + 1 : deltaOre + 1);
                            suffData = deltaOre < 0 ? suffissoDataPrec : suffissoData;
                            suffOra = "H" + ora;
                        }
                        siglaInformazione = Regex.Replace(siglaInformazione, @"\[[-+]?\d+\]", "");
                    }
                    else if (oraLock > 0)
                    {
                        suffData = suffissoData;
                        suffOra = "H" + oraLock;

                    }
                    else
                    {
                        if (suffissoData == "DATA1")
                        {
                            int ora = tmpdeltaNeg == 0 ? 1 : Math.Abs(tmpdeltaNeg);// +(info["Data0H24"].Equals("1") ? 0 : 1);
                            suffData = suffissoData;
                            suffOra = "H" + ora;
                        }
                        else
                        {
                            suffData = suffissoData;
                            suffOra = "H1";
                        }
                    }
                    Range rng = _definedNames.Get(siglaEntita, siglaInformazione, suffData, suffOra);
                    rng.Lock = oraLock > 0;
                    

                    return rng.ToString();
                }, RegexOptions.IgnoreCase);
                return formula;
            }
            deltaNeg = 0;
            deltaPos = 0;

            return "";
        }
        /// <summary>
        /// Launcher per l'aggiornamento dei dati.
        /// </summary>
        public override void UpdateData()
        {
            SplashScreen.UpdateStatus("Aggiorno informazioni");
            CancellaDati();
            AggiornaDateTitoli();
            CaricaParametri();
            CaricaInformazioni();
            AggiornaGrafici();
            SplashScreen.UpdateStatus("Aggiorno colori date");
            UpdateDayColor();
            MakeCellsDisabled();
        }
        #region UpdateData

        /// <summary>
        /// Cancella le informazioni da aggiornare in tutti i giorni.
        /// </summary>
        private void CancellaDati()
        {
            CancellaDati(Workbook.DataAttiva, true);
        }
        /// <summary>
        /// Cancella le informazioni a partire da giorno. Se all è a true, cancella tutti giorni successivi altrimenti cancella il solo giorno.
        /// </summary>
        /// <param name="giorno">Data di partenza della cancellazione.</param>
        /// <param name="all">Se true cancella tutti i dati a partire dalla data di partenza.</param>
        private void CancellaDati(DateTime giorno, bool all = false)
        {
            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND IdApplicazione = " + Workbook.IdApplicazione; // AND (Gerarchia = '' OR Gerarchia IS NULL )";

            string suffissoData = Date.GetSuffissoData(giorno);
            int colOffset = _definedNames.GetColOffset();
            if (!all)
                colOffset = Date.GetOreGiorno(giorno);

            foreach (DataRowView entita in categoriaEntita)
            {
                SplashScreen.UpdateStatus("Cancello dati " + entita["DesEntita"]);
                DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND FormulaInCella = '0' AND SiglaTipologiaInformazione NOT LIKE 'TITOLO%' AND IdApplicazione = " + Workbook.IdApplicazione;// AND ValoreDefault IS NULL";

                foreach (DataRowView info in informazioni)
                {
                    int col = all ? _definedNames.GetFirstCol() : _definedNames.GetColFromDate(suffissoData);
                    object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                    
                    if (Struct.tipoVisualizzazione == "O")
                    {
                        int row = _definedNames.GetRowByName(siglaEntita, info["SiglaInformazione"]);
                        CancellaInformazione(info, row, col, colOffset);
                    }
                    else
                    {
                        DateTime dataInizio = giorno;
                        DateTime dataFine = giorno;
                        if(all)
                        {
                            dataInizio = Workbook.DataAttiva;
                            dataFine = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni);
                        }

                        CicloGiorni(dataInizio, dataFine, (oreGiorno, suffData, g) =>
                        {
                            SplashScreen.UpdateStatus("Cancello dati " + g.ToShortDateString());

                            int row = _definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], suffData);
                            CancellaInformazione(info, row, col, oreGiorno);
                        });
                    }
                }
                //reset colonna 24esima 25esima ora
                if (all && (Struct.tipoVisualizzazione == "V" || Struct.tipoVisualizzazione == "R") && informazioni.Count > 0)
                {
                    DateTime dataInizio = Workbook.DataAttiva;
                    DateTime dataFine = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni);

                    object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];

                    CicloGiorni(dataInizio, dataFine, (oreGiorno, suffData, g) =>
                    {
                        Range rngData = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], suffData), _definedNames.GetFirstCol(), informazioni.Count, oreGiorno);                        

                        int ore = Date.GetOreGiorno(g);
                        if (ore == 23) 
                        {
                            _ws.Range[rngData.Columns[rngData.Columns.Count - 2, rngData.Columns.Count - 1].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                        }
                        else if (ore == 24)
                        {
                            _ws.Range[rngData.Columns[rngData.Columns.Count - 2].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
                            _ws.Range[rngData.Columns[rngData.Columns.Count - 1].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                        }
                        else if (ore == 25)
                        {
                            _ws.Range[rngData.Columns[rngData.Columns.Count - 2].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
                            _ws.Range[rngData.Columns[rngData.Columns.Count - 1].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
                        }
                    });
                }
            }
        }
        /// <summary>
        /// Cancella il contenuto di una specifica riga di informazione.
        /// </summary>
        /// <param name="info">Informazione da cancellare</param>
        /// <param name="row">Riga dell'informazione</param>
        /// <param name="col">Colonna di inizio dell'informazione</param>
        /// <param name="realColOffset">Offset di colonne reale.</param>
        private void CancellaInformazione(DataRowView info, int row, int col, int realColOffset)
        {
            if (info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
            {
                Excel.Range rngData = _ws.Range[Range.GetRange(row, col - 1)];
                rngData.Value = "";
            }
            else
            {
                if (_struttura.visData0H24 && info["Data0H24"].Equals("0"))
                {
                    col++;
                    realColOffset--;
                }
                Excel.Range rngData = _ws.Range[Range.GetRange(row, col, 1, realColOffset)];
                rngData.Value = "";

                rngData.ClearComments();
                Style.RangeStyle(rngData, backColor: info["BackColor"], foreColor: info["ForeColor"]);
            }
        }
        /// <summary>
        /// Aggiorna le date dei titolo (per il caso in cui l'aggiornamento venga da un cambio giorno).
        /// </summary>
        public override void AggiornaDateTitoli()
        {
            if (Struct.tipoVisualizzazione == "O")
            {
                int row = _struttura.rigaBlock - 2;
                for (int i = 0; i < _definedNames.DaySuffx.Length; i++)
                {
                    if (_definedNames.DaySuffx[i] != "DATA0")
                    {
                        int col = _definedNames.GetColFromDate(_definedNames.DaySuffx[i]);
                        _ws.Range[Range.GetRange(row, col)].Value = Date.GetDataFromSuffisso(_definedNames.DaySuffx[i]);
                    }
                }
            }
            else if (Struct.tipoVisualizzazione == "V")
            {
                DefinedNames gotos = new DefinedNames(_ws.Name, DefinedNames.InitType.GOTOsThisSheet);

                for (int i = 0; i <= Struct.intervalloGiorni; i++)
                {
                    DateTime giorno = Workbook.DataAttiva.AddDays(i);
                    string suffissoData = Date.GetSuffissoData(giorno);
                    
                    int row = _definedNames.GetRowByName(suffissoData, "T");
                    int col = _definedNames.GetFirstCol();
                    _ws.Range[Range.GetRange(row, col)].Value = giorno;

                    row += 2;
                    col -= (_visSelezione + 1);
                    if (_ws.Range[Range.GetRange(row, col)].Value != null)
                        _ws.Range[Range.GetRange(row, col)].Value = giorno;

                    _ws.Range[gotos.GetFromAddressGOTO(i)].Value = giorno;

                }
            }
            else if (Struct.tipoVisualizzazione == "R")
            {
                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
                
                categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL) AND IdApplicazione = " + Workbook.IdApplicazione;
                
                foreach (DataRowView entita in categoriaEntita)
                {
                    informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                    foreach (DataRowView info in informazioni)
                    {
                        CicloGiorni(Workbook.DataAttiva, Workbook.DataAttiva.AddDays(Struct.intervalloGiorni), (oreGiorno, suffissoData, giorno) =>
                        {
                            int row = _definedNames.GetRowByNameSuffissoData(info["SiglaEntita"], info["SiglaInformazione"], suffissoData);
                            int col = _definedNames.GetFirstCol() - _visSelezione;
                            _ws.Range[Range.GetRange(row, col)].Value = giorno;
                        });
                    }
                }
            }
        }
        /// <summary>
        /// Carica i parametri e valori di default.
        /// </summary>
        protected void CaricaParametri()
        {
            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL ) AND IdApplicazione = " + Workbook.IdApplicazione;
            _dataInizio = Workbook.DataAttiva;

            foreach (DataRowView entita in categoriaEntita)
            {
                SplashScreen.UpdateStatus("Carico parametri " + entita["DesEntita"]);
                _dataFine = _dataFineUP[entita["SiglaEntita"]];

                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaTipologiaParametro IS NOT NULL AND IdApplicazione = " + Workbook.IdApplicazione;
                InsertParametri();

                SplashScreen.UpdateStatus("Aggiorno valori di default " + entita["DesEntita"]);
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND ValoreDefault IS NOT NULL AND IdApplicazione = " + Workbook.IdApplicazione;
                InsertFormuleValoriDefault();
            }
        }
        /// <summary>
        /// Aggiorna la colorazione della barra superiore della data (o delle varie barre se in visualizzazione verticale) e delle celle GOTO (solo in visualizzazione verticale) in base allo schema colori basato sui giorni.
        /// </summary>
        public void UpdateDayColor()
        {
            ColoraDataOra();
            ColoraGOTO();
        }

        #endregion  
      
        #region Disabilita celle per mercati MB/MI
        
        /// <summary>
        /// Applica un pattern alle informazioni nella parte che non è editabile a causa della chiusura del mercato.
        /// </summary>
        public override void MakeCellsDisabled()
        {
            if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1") || Workbook.IdApplicazione == 18)
            {
                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND IdApplicazione = " + Workbook.IdApplicazione; // AND (Gerarchia = '' OR Gerarchia IS NULL )";

                int colOffset = _definedNames.GetColOffset();
                int marketOffset = 0;
                if (Workbook.IdApplicazione == 18)
                {
                    marketOffset = Simboli.GetMarketOffsetMI(Workbook.Mercato, Workbook.DataAttiva);
                }
                else
                {
                    marketOffset = Simboli.GetMarketOffset(DateTime.Now.Hour);
                }

                foreach (DataRowView entita in categoriaEntita)
                {
                    DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
                    informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaTipologiaInformazione NOT LIKE 'TITOLO%' AND SiglaTipologiaInformazione <> 'CHECK' AND IdApplicazione = " + Workbook.IdApplicazione;// AND ValoreDefault IS NULL";

                    int col = _definedNames.GetFirstCol();
                    //int col = _definedNames.GetColData1H1();
                    foreach (DataRowView info in informazioni)
                    {
                        //06/02/2017 MOD: shift se visualizzo DATA0.H24
                        

                        object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                        if (Struct.tipoVisualizzazione == "O")
                        {
                            int row = _definedNames.GetRowByName(siglaEntita, info["SiglaInformazione"]);
                            MakeCellsDisabled(info, row, col, colOffset, marketOffset);
                        }
                        else
                        {
                            CicloGiorni(Workbook.DataAttiva, Workbook.DataAttiva.AddDays(Struct.intervalloGiorni), (oreGiorno, suffData, giorno) =>
                            {
                                int row = _definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], suffData);
                                MakeCellsDisabled(info, row, col, Date.GetOreGiorno(giorno), marketOffset);
                            });
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Applica un pattern alla specifica riga di informazione nella parte che non è editabile a causa della chiusura del mercato.
        /// </summary>
        /// <param name="info">Informazione.</param>
        /// <param name="row">Riga dell'informazione.</param>
        /// <param name="col">Colonna di inizio dell'informazione.</param>
        /// <param name="colOffset">Offset dell'intera riga per fare il clean da situazioni precedenti.</param>
        /// <param name="disabledOffset">Offset a cui applicare il pattern.</param>
        private void MakeCellsDisabled(DataRowView info, int row, int col, int colOffset, int disabledOffset)
        {
            //clear del pattern
            Range rng = new Range(row, col, 1, colOffset);
            Style.RangeStyle(_ws.Range[rng.ToString()], pattern: Excel.XlPattern.xlPatternAutomatic);

            //applico stile
            rng = new Range(row, col, 1, disabledOffset);
            Style.RangeStyle(_ws.Range[rng.ToString()], pattern: Excel.XlPattern.xlPatternGray50);
        }
        
        #endregion



        public void Dispose()
        {
            if (!_disposed)
            {
                GC.SuppressFinalize(this);
                _disposed = true;
            }
        }

        #endregion
    }
}
