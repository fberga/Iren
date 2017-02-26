using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace Iren.PSO.Base
{
    public class DefinedNames
    {
        #region Variabili

        string _sheet;
        List<string> _days;

        InitType _initType;

        protected Dictionary<string, int> _defDatesIndexByName = new Dictionary<string,int>();
        protected Dictionary<int, string> _defDatesIndexByCol = new Dictionary<int, string>();

        protected Dictionary<string, int> _defNamesIndexByName = new Dictionary<string, int>();
        protected ILookup<int, string> _defNamesIndexByRow;

        protected Dictionary<string, object> _addressFrom = new Dictionary<string, object>();
        protected Dictionary<object, string> _addressTo = new Dictionary<object, string>();

        protected Dictionary<int, string> _editable = new Dictionary<int, string>();
        protected List<int> _saveDB = new List<int>();
        protected List<int> _toNote = new List<int>();

        protected List<CheckObj> _check = new List<CheckObj>();
        protected List<Selection> _selections = new List<Selection>();

        /// <summary>
        /// Specifica le varie tipologie di inizializzazione utilizzate.
        /// </summary>
        public enum InitType
        {
            /// <summary>
            /// Inizializza l'oggetto con tutte le funzionalità annesse.
            /// </summary>
            All, 
            /// <summary>
            /// Inizializza l'oggetto con la struttura dei nomi e dei campi per le selection.
            /// </summary>
            Naming, 
            /// <summary>
            /// Inizializza l'oggetto con la sola struttura dei GOTO di tutti i fogli.
            /// </summary>
            GOTOs, 
            /// <summary>
            /// Inizializza l'oggetto con i GOTO di un solo foglio (quello utilizzato per inizializzare l'oggetto).
            /// </summary>
            GOTOsThisSheet, 
            /// <summary>
            /// Inizializza l'oggetto con la sola struttura dei campi editabili.
            /// </summary>
            Editable, 
            /// <summary>
            /// Inizializza l'oggetto con la sola struttura dei campi da salvare sul database in seguito a modifica.
            /// </summary>
            SaveDB, 
            /// <summary>
            /// Inizializza l'oggetto con la struttura dei check e, se presenti, anche quella dei nomi.
            /// </summary>
            CheckNaming, 
            /// <summary>
            /// Inizializza l'oggetto con la sola struttura dei check.
            /// </summary>
            Check, 
            /// <summary>
            /// Inizializza l'oggetto con la sola struttura dei campi per le selezioni.
            /// </summary>
            Selection
        }

        #endregion

        #region Proprietà

        /// <summary>
        /// Restituisce tutti i suffissi dei giorni definiti nella struttura. Ad esempio se l'intervallo dei giorni è 2, restituirà una lista composta da { "DATA1", "DATA2" }. (Deve essere inizializzato con Naming, All)
        /// </summary>
        public string[] DaySuffx
        {
            get
            {
                if (_initType == InitType.CheckNaming || _initType == InitType.Naming || _initType == InitType.All)
                    return _days.ToArray();
                else
                    throw new MemberAccessException("L'oggetto non è stato inizializzato in modo da poter utilizzare questa risorsa. Inizializzarlo con tipologia Naming, CheckNaming oppure All.");
            }
        }
        /// <summary>
        /// Restituisce il nome del foglio a cui i nomi definiti fanno riferimento. (Read Only)
        /// </summary>
        public string Sheet
        {
            get { return _sheet; }
        }
        /// <summary>
        /// Restituisce la lista di di range editabili: sono indicizzati per riga e contengono la stringa di indirizzo del range. (Deve essere inizializzato con Editable, All)
        /// </summary>
        public Dictionary<int, string> Editable
        {
            get 
            {
                if (_initType == InitType.Editable || _initType == InitType.All)
                    return _editable; 
                else
                    throw new MemberAccessException("L'oggetto non è stato inizializzato in modo da poter utilizzare questa risorsa. Inizializzarlo con tipologia Editable oppure All.");
            }
        }
        /// <summary>
        /// Verifica se è presente la DATA0H24 ovvero l'ora 24 del giorno prima di quello di inizio. (Deve essere inizializzato con Naming, All, C)
        /// </summary>
        public bool HasData0H24
        {
            //se è definita la colonna DATA0H24, sarà la prima.
            get 
            {
                if (_initType == InitType.CheckNaming || _initType == InitType.Naming || _initType == InitType.All)
                    return _defDatesIndexByName.First().Key == GetName(Date.GetSuffissoData(Workbook.DataAttiva.AddDays(-1)), Date.GetSuffissoOra(24)); 
                else
                    throw new MemberAccessException("L'oggetto non è stato inizializzato in modo da poter utilizzare questa risorsa. Inizializzarlo con tipologia Naming, CheckNaming, oppure All.");
            }
        }
        /// <summary>
        /// Restituisce la lista di tutti i check definiti (Deve essere inizializzato con Check, CheckNaming, All)
        /// </summary>
        public List<CheckObj> Checks
        {
            get 
            {
                if (_initType == InitType.CheckNaming || _initType == InitType.Check || _initType == InitType.All)
                    return _check;
                else
                    throw new MemberAccessException("L'oggetto non è stato inizializzato in modo da poter utilizzare questa risorsa. Inizializzarlo con tipologia Check, CheckNaming, oppure All.");
            }
        }

        #endregion

        #region Costruttori

        /// <summary>
        /// Inizializza la struttura di indicizzazione con i nomi.
        /// </summary>
        private void InitNaming()
        {
            DataTable definedNames = Workbook.Repository[DataBase.TAB.NOMI_DEFINITI];
            DataTable definedDates = Workbook.Repository[DataBase.TAB.DATE_DEFINITE];

            IEnumerable<DataRow> names =
                from DataRow r in definedNames.AsEnumerable()
                where r["Sheet"].Equals(_sheet)
                select r;

            _defNamesIndexByName = names.ToDictionary(r => r["Name"].ToString(), r => (int)r["Row"]);
            _defNamesIndexByRow = names.ToLookup(r => (int)r["Row"], r => r["Name"].ToString());

            IEnumerable<DataRow> dates =
                from DataRow r in definedDates.AsEnumerable()
                where r["Sheet"].Equals(_sheet)
                select r;

            DataView distinctDays = new DataView(definedDates);
            distinctDays.RowFilter = "Sheet = '" + _sheet + "'";
            _days =
                (from r in distinctDays.ToTable(true, "Date").AsEnumerable()
                 select r["Date"].ToString()).ToList();

            _defDatesIndexByName = dates.ToDictionary(r => GetName(r["Date"].ToString(), r["Hour"].ToString()), r => (int)r["Column"]);
            _defDatesIndexByCol = dates.ToDictionary(r => (int)r["Column"], r => GetName(r["Date"].ToString(), r["Hour"].ToString()));

        }
        /// <summary>
        /// Inizializza la struttura per i GOTO.
        /// </summary>
        /// <param name="thisSheet">Se true limita l'azione alla sola sheet corrente.</param>
        private void InitGOTOs(bool thisSheet = false)
        {
            DataTable addressFromTable = Workbook.Repository[DataBase.TAB.ADDRESS_FROM];
            DataTable addressToTable = Workbook.Repository[DataBase.TAB.ADDRESS_TO];

            _addressFrom =
               (from DataRow r in addressFromTable.AsEnumerable()
                where !thisSheet || r["Sheet"].Equals(_sheet)
                select r).ToDictionary(
                    r => r["AddressFrom"].ToString(),
                    r => r["SiglaEntita"]
                );

            _addressTo =
               (from DataRow r in addressToTable.AsEnumerable()
                where !thisSheet || r["Sheet"].Equals(_sheet)
                select r).ToDictionary(
                    r => r["SiglaEntita"],
                    r => r["AddressTo"].ToString()
                );
        }
        /// <summary>
        /// Inizializza la struttura per riconoscere le celle editabili.
        /// </summary>
        private void InitEditable()
        {
            DataTable editabili = Workbook.Repository[DataBase.TAB.EDITABILI];

            _editable =
                (from r in editabili.AsEnumerable()
                 where r["Sheet"].Equals(_sheet)
                 select r).ToDictionary(r => (int)r["Row"], r => r["Range"].ToString());
        }
        /// <summary>
        /// Inizialissa la struttura per riconoscere le celle da salvare sul DB.
        /// </summary>
        private void InitSaveDB()
        {
            DataTable saveDB = Workbook.Repository[DataBase.TAB.SALVADB];

            _saveDB =
                (from r in saveDB.AsEnumerable()
                 where r["Sheet"].Equals(_sheet)
                 select (int)r["Row"]).ToList();
        }
        /// <summary>
        /// Inizializza la struttura per riconoscere le celle su cui apporre commenti.
        /// </summary>
        private void InitToNote()
        {
            DataTable toNote = Workbook.Repository[DataBase.TAB.ANNOTA];

            _toNote =
                (from r in toNote.AsEnumerable()
                 where r["Sheet"].Equals(_sheet)
                 select (int)r["Row"]).ToList();
        }
        /// <summary>
        /// Inizializza la struttura per riconoscere le celle di check.
        /// </summary>
        private void InitCheck()
        {
            DataTable check = Workbook.Repository[DataBase.TAB.CHECK];

            _check =
                (from r in check.AsEnumerable()
                 where r["Sheet"].Equals(_sheet)
                 select new CheckObj(r["SiglaEntita"].ToString(), (string)r["Range"], (int)r["Type"])).ToList();
        }
        /// <summary>
        /// inizializza la struttura per riconoscere le righe che sono parte di una selezione.
        /// </summary>
        private void InitSelection()
        {
            DataTable selection = Workbook.Repository[DataBase.TAB.SELECTION];

            var groupings =
                (from r in selection.AsEnumerable()
                 where r["Sheet"].Equals(_sheet)
                 group r by r["Rif"] into g
                 select g);

            foreach (IGrouping<object, DataRow> g in groupings)
            {
                string rif = g.Key.ToString();
                Dictionary<string, int> peers = new Dictionary<string, int>();
                foreach (DataRow r in g)
                {
                    peers.Add((string)r["Range"], (int)r["Value"]);
                }
                _selections.Add(new Selection(rif, peers));
            }
        }

        public DefinedNames() { }
        public DefinedNames(string sheet, InitType type = InitType.Naming)
        {
            _sheet = sheet;
            _initType = type;
            switch (type)
            {
                case InitType.All:
                    InitNaming();
                    InitGOTOs();
                    InitEditable();
                    InitSaveDB();
                    InitCheck();
                    InitSelection();
                    break;
                case InitType.Naming:
                    InitNaming();
                    InitSelection();
                    InitGOTOs(true);
                    break;
                case InitType.GOTOs:
                    InitGOTOs();
                    break;
                case InitType.GOTOsThisSheet:
                    InitGOTOs(true);
                    break;
                case InitType.Editable:
                    InitEditable();
                    break;
                case InitType.SaveDB:
                    InitNaming();
                    InitSaveDB();
                    InitToNote();
                    break;
                case InitType.CheckNaming:
                    InitCheck();
                    if(_check.Count > 0)
                        InitNaming();
                    break;
                case InitType.Check:
                    InitCheck();
                    break;
                case InitType.Selection:
                    InitSelection();
                    break;
            }
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Inizializza le colonne "in un'unica soluzione". Calcola il numero di ore nell'intervallo di giorni e a partire dalla colonna di inizio genera tutti i riferimenti DATAORA-COLONNA. (Vale solo per i "fogli normali")
        /// </summary>
        /// <param name="dataInizio">Data iniziale dell'intervallo.</param>
        /// <param name="dataFine">Data finale dell'intervallo.</param>
        /// <param name="colStart">Prima colonna da inizializzare</param>
        /// <param name="data0H24">Indica se esiste o no la DATA0H24</param>
        public void DefineDates(DateTime dataInizio, DateTime dataFine, int colStart, bool data0H24)
        {
            if (data0H24)
            {
                string data = GetName(Date.GetSuffissoData(dataInizio.AddDays(-1)), Date.GetSuffissoOra(24));
                _defDatesIndexByName.Add(data, colStart);
                _defDatesIndexByCol.Add(colStart, data);
                colStart++;
            }

            for (DateTime giorno = dataInizio; giorno <= dataFine; giorno = giorno.AddDays(1))
            {
                int oreGiorno = Struct.tipoVisualizzazione == "O" ? Date.GetOreGiorno(giorno) : 25;

                string suffissoData = Date.GetSuffissoData(giorno);
                for (int ora = 0; ora < oreGiorno; ora++)
                {
                    string data = GetName(suffissoData, Date.GetSuffissoOra(ora + 1));
                    _defDatesIndexByName.Add(data, colStart);
                    _defDatesIndexByCol.Add(colStart, data);
                    colStart++;
                }
                _days.Add(suffissoData);
            }
        }
        /// <summary>
        /// Collega il nome, costituito dall'insieme delle componenti in parts, con la riga in input.
        /// </summary>
        /// <param name="riga">Riga da collegare.</param>
        /// <param name="parts">Lista delle componenti del nome.</param>
        public void AddName(int riga, params object[] parts)
        {
            _defNamesIndexByName.Add(GetName(parts), riga);
            //_defNamesIndexByRow(riga, GetName(parts));
        }
        /// <summary>
        /// Collega il nome, costituito dall'insieme delle componenti in parts, con la colonna in input. (Utilizzato nelle customizzazioni dei fogli e nel riepilogo)
        /// </summary>
        /// <param name="col">Colonna da collegare.</param>
        /// <param name="parts">Lista delle componenti del nome.</param>
        public void AddCol(int col, params object[] parts)
        {
            _defDatesIndexByName.Add(GetName(parts), col);
            _defDatesIndexByCol.Add(col, GetName(parts));
        }
        /// <summary>
        /// Collega l'entità alla cella GOTO dove è posizionato il tasto da cliccare.
        /// </summary>
        /// <param name="siglaEntita">L'entità a cui si riferisce il goto.</param>
        /// <param name="addressFrom">L'indirizzo deve è posizionato il tasto.</param>
        public void AddGOTO(object siglaEntita, string addressFrom)
        {
            _addressFrom.Add("'" + _sheet + "'!" + addressFrom, siglaEntita);
        }
        /// <summary>
        /// Collega l'entita alla cella GOTO del tasto e alla cella da richiamare quando si clicca il tasto.
        /// </summary>
        /// <param name="siglaEntita">L'entità a cui si riferisce il goto.</param>
        /// <param name="addressFrom">L'indirizzo deve è posizionato il tasto.</param>
        /// <param name="addressTo">L'indirizzo a cui punta l'azione del GOTO.</param>
        public void AddGOTO(object siglaEntita, string addressFrom, string addressTo)
        {
            AddGOTO(siglaEntita, addressFrom);
            _addressTo.Add(siglaEntita, "'" + _sheet + "'!" + addressTo);
        }
        /// <summary>
        /// Nel caso in cui non sia stato assegnato un indirizzo di destinazione al GOTO, collega all'entità questo indirizzo.
        /// </summary>
        /// <param name="siglaEntita">Entità a cui collegare il GOTO.</param>
        /// <param name="addressTo">Indirizzo di arrivo dell'azione.</param>
        public void ChangeGOTOAddressTo(object siglaEntita, string addressTo)
        {
            _addressTo[siglaEntita] = "'" + _sheet + "'!" + addressTo;
        }
        /// <summary>
        /// Marca la il range come editabile suddividendo il tutto per righe.
        /// </summary>
        /// <param name="row">Riga a cui si riferisce il range.</param>
        /// <param name="rng">Range editabile.</param>
        public void SetEditable(int row, Range rng)
        {
            if (!_editable.ContainsKey(row))
                _editable.Add(row, rng.ToString());
            else
                _editable[row] += "," + rng.ToString();
        }
        /// <summary>
        /// Marca l'insieme di celle come appartenenti ad una selezione.
        /// </summary>
        /// <param name="rif">Cella di riferimento dove scrivere il valore di selezione</param>
        /// <param name="peers">Celle in cui cliccare per cambiare la selezione</param>
        public void SetSelection(string rif, Dictionary<string, int> peers)
        {
            _selections.Add(new Selection(rif, peers));
        }
        /// <summary>
        /// Marca la riga come da salvare sul database.
        /// </summary>
        /// <param name="row">Riga da salvare.</param>
        public void SetSaveDB(int row)
        {
            if (!_saveDB.Contains(row))
                _saveDB.Add(row);
        }
        /// <summary>
        /// Marca la riga come da annotare (ovvero su cui verrà aggiunta la nota da segnalare all'utente) sul database.
        /// </summary>
        /// <param name="row">Riga da annotare.</param>
        public void SetToNote(int row)
        {
            if (!_toNote.Contains(row))
                _toNote.Add(row);
        }

        

        /// <summary>
        /// Marca la riga come check.
        /// </summary>
        /// <param name="siglaEntita">Entità a cui appartiene il check.</param>
        /// <param name="range">Range delle celle di check.</param>
        /// <param name="type">Tipo di check (estratto dal DB).</param>
        public void AddCheck(string siglaEntita, string range, int type)
        {            
                _check.Add(new CheckObj(siglaEntita, range, type));
        }
        /// <summary>
        /// Verifica se la riga sia da salvare sul Database.
        /// </summary>
        /// <param name="row">Riga da verificare.</param>
        /// <returns>True se la riga è da salvare, False altrimenti.</returns>
        public bool SaveDB(int row)
        {
            return _saveDB.Contains(row);
        }
        /// <summary>
        /// Verifica se la riga sia da annotare o no.
        /// </summary>
        /// <param name="row">Riga da verificare.</param>
        /// <returns>True se la riga è da annotare, False altrimenti</returns>
        public bool ToNote(int row)
        {
            return _toNote.Contains(row);
        }
        /// <summary>
        /// Restituisce la prima colonna definita. Solitamente coinciderà con la colonna "colBlock" definita nella struttura del foglio.
        /// </summary>
        /// <returns>L'indirizzo della prima colonna definita.</returns>
        public int GetFirstCol()
        {
            return _defDatesIndexByCol.ElementAt(0).Key;
        }

        public int GetColData1H1()
        {
            return _defDatesIndexByName[GetName(Date.SuffissoDATA1, Date.GetSuffissoOra(1))];
        }

        /// <summary>
        /// Restituisce la prima riga definita. Solitamente coinciderà con la riga "rowBlock" definita nella struttura del foglio.
        /// </summary>
        /// <returns>L'indirizzo della prima riga definita.</returns>
        public int GetFirstRow()
        {
            return _defNamesIndexByName.ElementAt(0).Value;
        }
        /// <summary>
        /// Restituisce l'indirizzo dell'ultima colonna definita.
        /// </summary>
        /// <returns>Indirizzo dell'ultima colonna definita.</returns>
        public int GetLastCol()
        {
            return _defDatesIndexByCol.Last().Key;
        }
        /// <summary>
        /// Restituisce l'indirizzo della colonna a partire dalla DataAttiva del foglio all'ora uno.
        /// </summary>
        /// <returns>L'indirizzo della colonna corrispondente a DATA1.H1.</returns>
        public int GetColFromDate()
        {
            return GetColFromDate(Date.SuffissoDATA1);
        }
        /// <summary>
        /// Restituisce l'indirizzo della colonna a partire da giorno del foglio all'ora uno. (Utilizzato nei fogli normali)
        /// </summary>
        /// <param name="giorno">Il giorno di cui trovare la colonna H1.</param>
        /// <returns>L'indirizzo della colonna corrispondente a SuffissoData(giorno).H1</returns>
        public int GetColFromDate(DateTime giorno)
        {
            return GetColFromDate(Date.GetSuffissoData(giorno));
        }
        /// <summary>
        /// Restituisce l'indirizzo della colonna a partire dal suffisso data e dal suffisso ora. (Utilizzato nei fogli normali)
        /// </summary>
        /// <param name="suffissoData">Suffisso data di cui trovare la colonna.</param>
        /// <param name="suffissoOra">Suffisso ora di cui trovare la colonna.</param>
        /// <returns>L'indirizzo della colonna suffissoData.suffissoOra.</returns>
        public int GetColFromDate(string suffissoData, string suffissoOra = "H1")
        {
            if (Struct.tipoVisualizzazione == "V" || Struct.tipoVisualizzazione == "R")
                suffissoData = Date.SuffissoDATA1;

            string name = GetName(suffissoData, suffissoOra);
            return _defDatesIndexByName[name];
        }
        /// <summary>
        /// Restituisce l'indirizzo della colonna a partire da un nome. (Utilizzato nel Riepilogo e fogli custom)
        /// </summary>
        /// <param name="parts">Parti che compongono il nome</param>
        /// <returns>L'indirizzo della colonna indicata dal nome.</returns>
        public int GetColFromName(params object[] parts)
        {
            return _defDatesIndexByName[GetName(parts)];
        }
        /// <summary>
        /// Restituisce il numero di colonne totali del Riepilogo.
        /// </summary>
        /// <returns>Restituisce il numero di colonne totali del Riepilogo.</returns>
        public int GetColOffsetRiepilogo()
        {
            return _defDatesIndexByName.Count;
        }
        /// <summary>
        /// Restituisce il numero totale de
        /// </summary>
        /// <returns></returns>
        public int GetRowOffset()
        {
            return _defNamesIndexByName.Count;
        }
        /// <summary>
        /// Restituisce il numero di colonne definite da utilizzare come offset sul foglio. Da utilizzare solo nei fogli in cui le colonne sono tutte contigue altrimenti il risultato non rappresenterà l'offset effettivo sul foglio.
        /// </summary>
        /// <returns>Il numero di colonne definite</returns>
        public int GetColOffset()
        {
            if (Struct.tipoVisualizzazione == "O")
                return _defDatesIndexByName.Count;

            return 25;
        }
        /// <summary>
        /// Restituisce il numero di colonne che vanno dalla data iniziale a quella passata come parametro. Può essere utilizzato in alternativa al metodo Date.GetDayOffset(dataInizio, dataFine). Da ricordare che questo metodo, dove presente, conteggia la colonna della DATA0H24.
        /// </summary>
        /// <param name="data">La data fino a cui conteggiare le colonne definite.</param>
        /// <returns>Il numero di colonne definite fino a data</returns>
        public int GetColOffset(DateTime data)
        {
            return GetColOffset(Date.GetSuffissoData(data));
        }
        /// <summary>
        /// Restituisce il numero di colonne che vanno dalla data iniziale a quella passata come parametro. Può essere utilizzato in alternativa al metodo Date.GetOreIntervallo(dataInizio, dataFine). Da ricordare che questo metodo, dove presente, conteggia la colonna della DATA0H24.
        /// </summary>
        /// <param name="suffissoData">Il suffisso della data fino a cui conteggiare le colonne definite.</param>
        /// <returns>Il numero di colonne definite fino a data</returns>
        public int GetColOffset(string suffissoData)
        {
            var date =
                from kv in _defDatesIndexByName
                where kv.Key.Substring(0, suffissoData.Length).CompareTo(suffissoData) <= 0
                select kv;

            return date.Count();
        }
        /// <summary>
        /// Restituisce il numero di colonne definite per il giorno indicato dal parametro suffissoData. Se la tipologia di visualizzazione è Verticale, il numero restituito sarà sempre 25. Può essere utilizzato in alternativa a Date.GetOreGiorno(suffissoData).
        /// </summary>
        /// <param name="suffissoData">Il suffisso della data di cui conteggiare le colonne.</param>
        /// <returns>Numero di colonne definite per il giorno.</returns>
        public int GetDayOffset(string suffissoData)
        {
            if (Struct.tipoVisualizzazione == "V")
                return 25;

            var date =
                from kv in _defDatesIndexByName
                where kv.Key.StartsWith(suffissoData)
                select kv;

            return date.Count();
        }
        /// <summary>
        /// Restituisce il numero di colonne definite per il giorno indicato dal parametro giorno. Se la tipologia di visualizzazione è Verticale, il numero restituito sarà sempre 25. Può essere utilizzato in alternativa a Date.GetOreGiorno(giorno).
        /// </summary>
        /// <param name="giorno">Data di cui conteggiare le colonne.</param>
        /// <returns>Numero di colonne definite per il giorno.</returns>
        public int GetDayOffset(DateTime giorno)
        {
            return GetDayOffset(Date.GetSuffissoData(giorno));
        }
        /// <summary>
        /// Restituisce la riga collegata al nome passato come parametro.
        /// </summary>
        /// <param name="parts">Le componenti del nome di cui cercare la riga.</param>
        /// <returns>La riga del nome passato come parametro.</returns>
        public int GetRowByName(params object[] parts)
        {
            return _defNamesIndexByName[GetName(parts)];
        }
        /// <summary>
        /// Restituisce la riga collegata al nome passato come parametro.
        /// </summary>
        /// <param name="name">Il nome di cui cercare la riga.</param>
        /// <returns>La riga del nome passato come parametro.</returns>
        public int GetRowByName(string name)
        {
            return _defNamesIndexByName[name];
        }
        /// <summary>
        /// Caso particolare di GetRowByName in cui si presuppone che il nome sia composto da sigleEntita, siglaInformazione e suffissoData
        /// </summary>
        /// <param name="siglaEntita">La sigla dell'entità.</param>
        /// <param name="siglaInformazione">La sigla dell'informasione.</param>
        /// <param name="suffissoData">Il suffisso della data.</param>
        /// <returns>La riga del nome passato come parametro.</returns>
        public int GetRowByNameSuffissoData(object siglaEntita, object siglaInformazione, string suffissoData)
        {
            string name = GetName(siglaEntita, siglaInformazione, Struct.tipoVisualizzazione == "O" ? "" : suffissoData);
            return GetRowByName(name);
        }
        /// <summary>
        /// Restituisce tutti i nomi definiti per quella riga.
        /// </summary>
        /// <param name="row">La riga dove sono definiti i nomi.</param>
        /// <returns>Lista dei nomi definiti nella riga row.</returns>
        public List<string> GetNameByRow(int row)
        {
            return _defNamesIndexByRow[row].ToList();
        }
        /// <summary>
        /// Restituisce la data definita per la colonna column.
        /// </summary>
        /// <param name="column">La colonna dove è definita la data.</param>
        /// <returns>La stringa che rappresenta la data nel formato SuffissoData.SuffissoOra.</returns>
        public string GetDateByCol(int column)
        {
            if (IsDataColumn(column))
                return _defDatesIndexByCol[column];
            else
                return Date.SuffissoDATA1;
        }
        /// <summary>
        /// Restituisce il nome definito nella cella indicata dall'indirizzo RC.
        /// </summary>
        /// <param name="row">Riga dell'indirizzo.</param>
        /// <param name="column">Colonna dell'indirizzo.</param>
        /// <returns>Restituisce il nome definito in quella cella.</returns>
        public string GetNameByAddress(int row, int column)
        {
            if(Struct.tipoVisualizzazione == "O")
                return GetName(GetNameByRow(row), GetDateByCol(column));

            string[] parts = GetDateByCol(column).Split(Simboli.UNION[0]);
            List<string> names = GetNameByRow(row);

            string name = IsDataColumn(column) ? GetNameByRow(row).First() : GetNameByRow(row).Last();
            
            if (parts.Length > 1)
                return GetName(name, parts.Last());

            return name;
        }
        /// <summary>
        /// Verifica se la colonna column fa parte del range dei dati oppure no. Ovvero maggiore di DATA1.H1 e minore di DATAn.H24.
        /// </summary>
        /// <param name="column">Colonna da verificare</param>
        /// <returns>True se è una colonna di dati, false altrimenti.</returns>
        public bool IsDataColumn(int column)
        {
            return column >= GetFirstCol() && column < GetFirstCol() + GetColOffset();
        }
        /// <summary>
        /// Verifica se il range passato contiene delle celle di check.
        /// </summary>
        /// <param name="rng">Range da verificare.</param>
        /// <returns>True se è un range contenente check, false altrimenti.</returns>
        public bool IsCheck(Range rng)
        {
            foreach (CheckObj chk in _check)
            {
                Range rngCheck = new Range(chk.Range);
                if (rngCheck.Contains(rng))
                    return true;
            }

            return false;
        }
        /// <summary>
        /// Verifica se il range passato fa parte di una selezione oppure no.
        /// </summary>
        /// <param name="rngPeer">Range da verificare.</param>
        /// <returns>True se il range è parte di una selezione, false altrimenti.</returns>
        public bool IsSelectionPeer(Range rngPeer)
        {
            foreach (Selection s in _selections)
                if (s.SelPeers.ContainsKey(rngPeer.ToString()))
                    return true;

            return false;
        }
        /// <summary>
        /// Restituisce l'oggetto di selezione e il valore corrispondente se il range in ingresso fa effettivamente parte di una selezione. Il range in ingresso fa parte di una delle celle di scelta della selezione e non la cella nascosta dove trascrivere il valore per la formattazione condizionale. Un valore resrtituito indica se la selezione è stata trovata.
        /// </summary>
        /// <param name="rngPeer">Il range che dovrebbe appartenere alla selezione.</param>
        /// <param name="sel">Se il range è parte di una selezione, viene restituito l'oggetto di selezione corrispondente che contiene tutti i riferimenti utili.</param>
        /// <param name="value">Se il range è parte di una selezione, viene anche restituito il valore corrispondente alla cella scelta.</param>
        /// <returns>True se viene effettivamente restituita una selezione, false altrimenti.</returns>
        public bool TryGetSelectionByPeer(Range rngPeer, out Selection sel, out int value)
        {
            foreach (Selection s in _selections)
            {
                if(s.SelPeers.ContainsKey(rngPeer.ToString()))
                {
                    sel = s;
                    value = s.SelPeers[rngPeer.ToString()];
                    return true;
                }
            }
            
            sel = null;
            value = -1;
            return false;
        }
        /// <summary>
        /// Restituisce l'oggetto di selezione se il range in ingresso fa effettivamente parte di una selezione. Il range in ingresso è la cella nascosta dove viene scritto il valore della selezione.
        /// </summary>
        /// <param name="rngRif">Il range che rappresenta la cella nascosta dove viene scritto il valore della selezione</param>
        /// <returns>L'oggetto di selezione corrispondente.</returns>
        public Selection GetSelectionByRif(Range rngRif)
        {
            foreach (Selection s in _selections)
            {
                Range rng = new Range(s.RifAddress);
                if (rng.Contains(rngRif))
                    return s;
            }
            return null;
        }
        /// <summary>
        /// Verifica se la riga sia o no collegata ad un nome definito.
        /// </summary>
        /// <param name="row">Riga da verificare</param>
        /// <returns>True se è un nome definito, false altrimenti.</returns>
        public bool IsDefined(int row)
        {
            return _defNamesIndexByRow.Contains(row);
        }
        /// <summary>
        /// Verifica se il nome è definito.
        /// </summary>
        /// <param name="parts">Le parti che compongono il nome</param>
        /// <returns>True se è definito, false altrimenti.</returns>
        public bool IsDefined(params object[] parts)
        {
            string name = GetName(parts);
            return _defNamesIndexByName.Count(kv => kv.Key.StartsWith(name)) > 0;
        }
        public bool IsDefinedExact(params object[] parts)
        {
            string name = GetName(parts);
            try
            {
                int row = _defNamesIndexByName[name];
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool IsEditable(int row)
        {
            return _editable.ContainsKey(row);
        }

        /// <summary>
        /// Metodo generico per la restituzione del range in base al nome. Lavora con il foglio di riepilogo come caso particolare. Per gli altri fogli, se il nome è composto da 2 parti, le considera SiglaEntita.SiglaInformazione. Se il nome è costituito da più parti, cerca il primo suffisso data valido e considera la parte antecedente come parte di riga mentre la successiva come parte di colonna. Se la parte di colonna non viene trovata, considera la colonna come la prima definita.
        /// </summary>
        /// <param name="parts">Le parti che compongono il nome.</param>
        /// <returns>Il range collegato al nome in ingresso</returns>
        public Range Get(params object[] parts)
        {
            if (_sheet == "Main")
            {
                int row = GetRowByName(GetName(parts[0]));
                
                int col = parts.Length == 3 ? GetColFromName(parts[1], parts[2]) : GetColFromName(parts[1]);

                return new Range(row, col);
            }
            else
            {
                if (parts.Length == 2)
                    return new Range(GetRowByName(GetName(parts)), GetFirstCol());

                List<string> nameParts = new List<string>();
                List<string> dateParts = new List<string>();
                bool date = false;
                foreach (var part in parts)
                {
                    date = date || Regex.IsMatch(part.ToString(), @"DATA\d+");
                    if (!date)
                        nameParts.Add(part.ToString());
                    else
                        dateParts.Add(part.ToString());
                }

                if (!date)
                    return new Range(GetRowByName(GetName(nameParts)), GetFirstCol());

                if (dateParts[0].Contains(Simboli.UNION))
                {
                    string[] suffissoDataOra = dateParts[0].Split(Simboli.UNION[0]);
                    dateParts = new List<string>() { suffissoDataOra[0], suffissoDataOra[1] };
                }
                else if (dateParts.Count == 1)
                    dateParts.Add(Date.GetSuffissoOra(1));

                int row = GetRowByName(GetName(nameParts, Struct.tipoVisualizzazione == "O" ? "" : dateParts[0]));
                int col = GetColFromDate(dateParts[0], dateParts[1]);
                
                return new Range(row, col);
            }
        }
        /// <summary>
        /// Metodo generico per la restituzione del range in base al nome. Lavora con il foglio di riepilogo come caso particolare. Per gli altri fogli, se il nome è composto da 2 parti, le considera SiglaEntita.SiglaInformazione. Se il nome è costituito da più parti, cerca il primo suffisso data valido e considera la parte antecedente come parte di riga mentre la successiva come parte di colonna. Se la parte di colonna non viene trovata, considera la colonna come la prima definita. Un valore indica se il range è stato trovato oppure no.
        /// </summary>
        /// <param name="rng">Il range collegato al nome in ingresso</param>
        /// <param name="parts">Le parti che compongono il nome.</param>
        /// <returns>True se il range è stato trovato, false altrimenti.</returns>
        public bool TryGet(out Range rng, params object[] parts)
        {
            try
            {
                rng = Get(parts);
                return true;
            }
            catch
            {
                rng = null;
                return false;
            }
        }
        /// <summary>
        /// Restituisce l'indirizzo a cui punta il GOTO a partire dalla cella premuta dall'utente.
        /// </summary>
        /// <param name="addressFrom">Indirizzo della cella premuta dall'utente.</param>
        /// <returns>Indirizzo di destinazione del GOTO.</returns>
        public string GetGotoFromAddress(string addressFrom)
        {
            if (_addressFrom.ContainsKey("'" + _sheet + "'!" + addressFrom))
                return GetGotoFromSiglaEntita(_addressFrom["'" + _sheet + "'!" + addressFrom]);

            return "";
        }
        /// <summary>
        /// Restituisce l'indirizzo a cui punta il GOTO a partire dalla sigla entità selezionata dall'utente.
        /// </summary>
        /// <param name="siglaEntita">Sigla entità della cella premuta dall'utente.</param>
        /// <returns>Indirizzo di destinazione del GOTO.</returns>
        public string GetGotoFromSiglaEntita(object siglaEntita)
        {
            if (_addressTo.ContainsKey(siglaEntita))
                return _addressTo[siglaEntita];

            return "";
        }
        /// <summary>
        /// Restituisce la lista di indirizzi delle celle che fungono da tasto per il GOTO a partire dalla sigla entità.
        /// </summary>
        /// <param name="siglaEntita">Sigla entità di cui trovare gli indirizzi di partenza</param>
        /// <returns>Lista di tutti gli indirizzi che fungono da tasto per il GOTO.</returns>
        public List<string> GetFromAddressGOTO(object siglaEntita)
        {
            List<string> o = 
                (from kv in _addressFrom
                where kv.Value.Equals(siglaEntita.ToString())
                select kv.Key).ToList();

            return o;
        }
        /// <summary>
        /// Restituisce la lista completa di indirizzi delle celle che fungono da tasto per i GOTO.
        /// </summary>
        /// <returns>Lista completa di indirizzi delle celle che fungono da tasto per i GOTO.</returns>
        public List<string> GetAllFromAddressGOTO()
        {
            List<string> o =
               (from kv in _addressFrom
                select kv.Key).ToList();

            return o;
        }
        /// <summary>
        /// Restituisce l'indirizzi della cella che funge da tasto per il GOTO in posizione i.
        /// </summary>
        /// <param name="i">Indice dell'elemento da cercare.</param>
        /// <returns>Indirizzo di partenza del GOTO.</returns>
        public string GetFromAddressGOTO(int i)
        {
            return _addressFrom.ElementAt(i).Key;
        }
        /// <summary>
        /// Verifica se nel foglio sono definiti dei check.
        /// </summary>
        /// <returns>True se nel foglio ci sono dei check, false altrimenti.</returns>
        public bool HasCheck()
        {
            return _check.Count > 0;
        }
        /// <summary>
        /// Verifica se nel foglio ci sono delle selezioni.
        /// </summary>
        /// <returns>True se nel foglio ci sono selezioni, false altrimenti.</returns>
        public bool HasSelections()
        {
            return _selections.Count > 0;
        }
        /// <summary>
        /// Verifica se nel foglio sono definiti dei nomi.
        /// </summary>
        /// <returns>True se ci sono dei nomi, false altrimenti.</returns>
        public bool HasNames()
        {
            return _defNamesIndexByName.Count > 0;
        }
        /// <summary>
        /// Salva l'intera struttura nel dataset locale.
        /// </summary>
        public void DumpToDataSet()
        {
            DataTable definedNames = Workbook.Repository[DataBase.TAB.NOMI_DEFINITI];
            DataTable definedDates = Workbook.Repository[DataBase.TAB.DATE_DEFINITE];
            DataTable addressFromTable = Workbook.Repository[DataBase.TAB.ADDRESS_FROM];
            DataTable addressToTable = Workbook.Repository[DataBase.TAB.ADDRESS_TO];
            DataTable editable = Workbook.Repository[DataBase.TAB.EDITABILI];
            DataTable saveDB = Workbook.Repository[DataBase.TAB.SALVADB];
            DataTable toNote = Workbook.Repository[DataBase.TAB.ANNOTA];
            DataTable check = Workbook.Repository[DataBase.TAB.CHECK];
            DataTable selection = Workbook.Repository[DataBase.TAB.SELECTION];

            ///////// nomi
            foreach (var ele in _defNamesIndexByName)
            {
                DataRow r = definedNames.NewRow();
                r["Sheet"] = _sheet;
                r["Name"] = ele.Key;
                r["Row"] = ele.Value;
                definedNames.Rows.Add(r);
            }

            ///////// date
            foreach (var ele in _defDatesIndexByName)
            {
                string[] dateTime = ele.Key.Split(Simboli.UNION[0]);

                DataRow r = definedDates.NewRow();
                r["Sheet"] = _sheet;
                r["Date"] = dateTime[0];
                r["Hour"] = dateTime.Length == 2 ? ele.Key.Split(Simboli.UNION[0])[1] : "";
                r["Column"] = ele.Value;
                definedDates.Rows.Add(r);
            }

            ///////// GOTO
            foreach (var ele in _addressFrom)
            {
                DataRow r = addressFromTable.NewRow();
                r["Sheet"] = _sheet;
                r["AddressFrom"] = ele.Key;
                r["SiglaEntita"] = ele.Value;
                addressFromTable.Rows.Add(r);
            }
            foreach (var ele in _addressTo)
            {
                DataRow r = addressToTable.NewRow();
                r["Sheet"] = _sheet;
                r["SiglaEntita"] = ele.Key;
                r["AddressTo"] = ele.Value;
                addressToTable.Rows.Add(r);
            }

            ///////// range editabili
            foreach (var ele in _editable)
            {
                DataRow r = editable.NewRow();
                r["Sheet"] = _sheet;
                r["Row"] = ele.Key;
                r["Range"] = ele.Value;
                editable.Rows.Add(r);
            }

            ///////// range da salvare sul db
            foreach (var ele in _saveDB)
            {
                DataRow r = saveDB.NewRow();
                r["Sheet"] = _sheet;
                r["Row"] = ele;
                saveDB.Rows.Add(r);
            }

            ///////// range che necessitano di commenti dopo modifica
            foreach (var ele in _toNote)
            {
                DataRow r = toNote.NewRow();
                r["Sheet"] = _sheet;
                r["Row"] = ele;
                toNote.Rows.Add(r);
            }

            ///////// celle con check
            foreach (var ele in _check)
            {
                DataRow r = check.NewRow();
                r["Sheet"] = _sheet;
                r["Range"] = ele.Range;
                r["SiglaEntita"] = ele.SiglaEntita;
                r["Type"] = ele.Type;
                check.Rows.Add(r);
            }

            ///////// celle che fanno parte di una selezione
            foreach (var ele in _selections)
            {
                foreach (var kv in ele.SelPeers)
                {
                    DataRow r = selection.NewRow();
                    r["Sheet"] = _sheet;
                    r["Rif"] = ele.RifAddress;
                    r["Range"] = kv.Key;
                    r["Value"] = kv.Value;
                    selection.Rows.Add(r);
                }
            }
        }

        #endregion

        #region Metodi Statici

        /// <summary>
        /// Restituisce il nome unito da Simboli.UNION dalle parti che lo compongono.
        /// </summary>
        /// <param name="parts">Lista di elementi</param>
        /// <param name="name">Ultima parte del nome</param>
        /// <returns>Stringa contenente il nome.</returns>
        public static string GetName(List<string> parts, string name)
        {            
            parts.Add(name);
            return GetName(parts);
        }
        /// <summary>
        /// Restituisce il nome unito da Simboli.UNION dalle parti che lo compongono.
        /// </summary>
        /// <param name="name">Prima parte del nome</param>
        /// <param name="parts">Lista di elementi</param>
        /// <returns>Stringa contenente il nome.</returns>
        public static string GetName(string name, List<string> parts)
        {
            List<string> list = new List<string>();
            list.Add(name);
            list.AddRange(parts);

            return GetName(list);
        }
        /// <summary>
        /// Restituisce il nome unito da Simboli.UNION dalle parti che lo compongono.
        /// </summary>
        /// <param name="parts">Array di liste di elementi</param>
        /// <returns>Stringa contenente il nome.</returns>
        public static string GetName(params List<string>[] parts)
        {
            string o = "";
            bool first = true;
            foreach (List<string> part in parts)
            {
                foreach (string part1 in part)
                {
                    if(part1 != null && part1 != "")
                    {
                        o += (!first ? Simboli.UNION : "") + part1;
                        first = false;
                    }
                }
            }
            return o;
        }
        /// <summary>
        /// Restituisce il nome unito da Simboli.UNION dalle parti che lo compongono.
        /// </summary>
        /// <param name="parts">Lista di oggetti che compongono il nome. Se sono oggetti validi si andrà a richiamare la funzione giusta tra gli overload.</param>
        /// <returns>Stringa contenente il nome.</returns>
        public static string GetName(params object[] parts)
        {
            List<string> list = new List<string>();
            foreach (object part in parts)
                if(part.GetType() == typeof(string))
                    list.Add(part.ToString());
                else if(part.GetType() == typeof(List<string>))
                {
                    foreach (var ele in (List<string>)part)
                        list.Add(ele);
                }

            return GetName(list);
        }
        /// <summary>
        /// Restituisce la struttura di default della tabella dei nomi. (NON MODIFICABILE SE NON CAMBIANDO TUTTO IL CODICE)
        /// </summary>
        /// <param name="name">Nome con cui inizializzare la tabella</param>
        /// <returns>Restituisce la tabella vuota.</returns>
        public static DataTable GetDefaultNameTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(String)},
                        {"Name", typeof(String)},
                        {"Row", typeof(int)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Name"] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Restituisce la struttura di default della tabella delle colonne (date per sheet normali, nomi per particolari). (NON MODIFICABILE SE NON CAMBIANDO TUTTO IL CODICE)
        /// </summary>
        /// <param name="name">Nome con cui inizializzare la tabella</param>
        /// <returns>Restituisce la tabella vuota.</returns>
        public static DataTable GetDefaultDateTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(String)},
                        {"Date", typeof(String)},
                        {"Hour", typeof(String)},
                        {"Column", typeof(int)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Date"], dt.Columns["Hour"] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Restituisce la struttura di default della tabella dei GOTO Address From. (NON MODIFICABILE SE NON CAMBIANDO TUTTO IL CODICE)
        /// </summary>
        /// <param name="name">Nome con cui inizializzare la tabella</param>
        /// <returns>Restituisce la tabella vuota.</returns>
        public static DataTable GetDefaultAddressFromTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(string)},
                        {"AddressFrom", typeof(string)},
                        {"SiglaEntita", typeof(string)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["AddressFrom"] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Restituisce la struttura di default della tabella dei GOTO Address To. (NON MODIFICABILE SE NON CAMBIANDO TUTTO IL CODICE)
        /// </summary>
        /// <param name="name">Nome con cui inizializzare la tabella</param>
        /// <returns>Restituisce la tabella vuota.</returns>
        public static DataTable GetDefaultAddressToTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(string)},
                        {"SiglaEntita", typeof(string)},
                        {"AddressTo", typeof(string)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["SiglaEntita"] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Restituisce la struttura di default della tabella dei campi editabili. (NON MODIFICABILE SE NON CAMBIANDO TUTTO IL CODICE)
        /// </summary>
        /// <param name="name">Nome con cui inizializzare la tabella</param>
        /// <returns>Restituisce la tabella vuota.</returns>
        public static DataTable GetDefaultEditableTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(string)},
                        {"Row", typeof(int)},
                        {"Range", typeof(string)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Row"] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Restituisce la struttura di default della tabella dei campi salvabili. (NON MODIFICABILE SE NON CAMBIANDO TUTTO IL CODICE)
        /// </summary>
        /// <param name="name">Nome con cui inizializzare la tabella</param>
        /// <returns>Restituisce la tabella vuota.</returns>
        public static DataTable GetDefaultSaveTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(string)},
                        {"Row", typeof(int)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Row"] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Restituisce la struttura di default della tabella dei campi da annotare. (NON MODIFICABILE SE NON CAMBIANDO TUTTO IL CODICE)
        /// </summary>
        /// <param name="name">Nome con cui inizializzare la tabella</param>
        /// <returns>Restituisce la tabella vuota.</returns>
        public static DataTable GetDefaultToNoteTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(string)},
                        {"Row", typeof(int)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Row"] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Restituisce la struttura di default della tabella dei campi di check. (NON MODIFICABILE SE NON CAMBIANDO TUTTO IL CODICE)
        /// </summary>
        /// <param name="name">Nome con cui inizializzare la tabella</param>
        /// <returns>Restituisce la tabella vuota.</returns>
        public static DataTable GetDefaultCheckTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(string)},
                        {"Range", typeof(string)},
                        {"SiglaEntita", typeof(string)},
                        {"Type", typeof(int)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Range"] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Restituisce la struttura di default della tabella dei campi selezione. (NON MODIFICABILE SE NON CAMBIANDO TUTTO IL CODICE)
        /// </summary>
        /// <param name="name">Nome con cui inizializzare la tabella</param>
        /// <returns>Restituisce la tabella vuota.</returns>
        public static DataTable GetDefaultSelectionTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(string)},
                        {"Rif", typeof(string)},
                        {"Range", typeof(string)},
                        {"Value", typeof(int)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Rif"], dt.Columns["Range"] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Restituisce il nome del foglio in base alla sigla entità in input.
        /// </summary>
        /// <param name="siglaEntita"></param>
        /// <returns>Nome del foglio che contiene l'entità in ingresso.</returns>
        public static string GetSheetName(object siglaEntita)
        {
            DataTable dt = Workbook.Repository[DataBase.TAB.NOMI_DEFINITI];

            List<Microsoft.Office.Interop.Excel.Worksheet> msdSheets = new List<Microsoft.Office.Interop.Excel.Worksheet>();

            foreach (var ws in Workbook.MSDSheets)
                msdSheets.Add(ws);


            string s =
                (from r in dt.AsEnumerable()
                 where r["Name"].ToString().Contains(siglaEntita.ToString()) && !r["Sheet"].Equals("Main") && !msdSheets.Contains(Workbook.Sheets[r["Sheet"]])
                 select r["Sheet"].ToString()).FirstOrDefault();

            return s ?? "";
        }

        #endregion
    }
}