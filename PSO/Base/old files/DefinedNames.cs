using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using Iren.ToolsExcel.Utility;
using System.Collections;

namespace Iren.ToolsExcel.Base
{
    public class DefinedNames
    {
        #region Costanti

        public struct Fields
        {
            public const string Foglio = "Foglio",
                Nome = "Nome",
                R1 = "R1",
                C1 = "C1",
                R2 = "R2",
                C2 = "C2",
                Editabile = "Editabile",
                SalvaDB = "SalvaDB",
                AnnotaModifica = "AnnotaModifica";
        }

        #endregion

        #region Variabili

        protected DataTable _definedNames;
        protected DataView _definedNamesView;
        protected string _foglio;
        
        #endregion

        #region Costruttori

        public DefinedNames(string foglio)
        {
            _foglio = foglio;
            _definedNames = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.NOMI_DEFINITI];
            _definedNamesView = new DataView(_definedNames);
        }

        #endregion

        #region Overload Operatori

        //public string this[params object[] parts] 
        //{
        //    get
        //    {
        //        return Get(parts);
        //    }
        //}
        //public string this[bool excludeDATA0H24, params object[] parts]
        //{
        //    get
        //    {
        //        return Get(excludeDATA0H24, parts);
        //    }
        //}
        
        
        public Tuple<int, int>[] this[params object[] parts] 
        {
            get 
            {
                return Get(parts);
            }
        }
        public Tuple<int, int>[] this[bool excludeDATA0H24, params object[] parts]
        {
            get
            {
                return Get(excludeDATA0H24, parts);
            }
        }
        
        //public Tuple<int, int>[] this[string key, bool excludeDATA0H24 = false]
        //{
        //    get
        //    {
        //        return Get(key, excludeDATA0H24);
        //    }
        //}

        public string[] this[int r1, int c1]
        {
            get
            {
                return Get(r1, c1);
            }
        }

        #endregion

        #region Metodi

        public void Add(string nome, Tuple<int, int> cella1, Tuple<int, int> cella2 = null, bool editabile = false, bool salvaDB = false, bool annotaModifica = false)
        {
            DataRow r = _definedNames.NewRow();
            cella2 = cella2 ?? cella1;
            r["Foglio"] = _foglio;
            r["Nome"] = nome;
            r["R1"] = cella1.Item1;
            r["C1"] = cella1.Item2;
            r["R2"] = cella2.Item1;
            r["C2"] = cella2.Item2;
            r["Editabile"] = editabile;
            r["SalvaDB"] = salvaDB;
            r["AnnotaModifica"] = annotaModifica;

            _definedNames.Rows.Add(r);
        }
        public void Add(string name, int row, int column, bool editabile = false, bool salvaDB = false, bool annotaModifica = false)
        {
            Add(name, Tuple.Create(row, column), editabile: editabile, salvaDB: salvaDB, annotaModifica: annotaModifica);
        }
        public void Add(string name, int row1, int column1, int row2, int column2, bool editabile = false, bool salvaDB = false, bool annotaModifica = false)
        {
            Add(name, Tuple.Create(row1, column1), Tuple.Create(row2, column2), editabile, salvaDB, annotaModifica);
        }

        public string[] Get(int row, int column)
        {
            string filter = "Foglio = '" + _foglio + "' AND R1 = " + row + " AND C1 = " + column;
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;
            
            if (_definedNamesView.Count == 0)
                return null;

            List<string> o = new List<string>();

            foreach (DataRowView name in _definedNamesView)
                o.Add(name["Nome"].ToString());

            return o.ToArray();
        }
        public Tuple<int, int>[] Get(bool excludeDATA0H24, params object[] parts)
        {
            if (excludeDATA0H24)
            {
                Array.Resize(ref parts, parts.Length + 2);
                parts[parts.Length - 2] = Simboli.EXCLUDE;
                parts[parts.Length - 1] = "DATA0.H24";
            }
            return Get(parts);
        }
        public Tuple<int, int>[] Get(params object[] parts)
        {
            if (parts.Length > 1)
            {
                int pos = Array.FindIndex(parts, ele => ele.ToString() == Simboli.EXCLUDE);
                pos = pos == -1 ? parts.Length : pos;

                string exclude = "";
                for (int i = pos + 1; i < parts.Length; i++)
                    exclude += " AND Nome NOT LIKE '%" + parts[i] + "%'";

                Array.Resize(ref parts, pos);

                if (Struct.tipoVisualizzazione == "V")
                {
                    string suffissoData = parts.Last().ToString().Contains("DATA") ? parts.Last().ToString() : parts[parts.Length - 2].ToString();
                    if (suffissoData.Contains("DATA"))
                    {
                        int oreGiorno = Date.GetOreGiorno(suffissoData);
                        if (oreGiorno == 23)
                        {
                            exclude += " AND Nome NOT LIKE '%H24'";
                            exclude += " AND Nome NOT LIKE '%H25'";
                        }
                        else if (oreGiorno == 24)
                            exclude += " AND Nome NOT LIKE '%H25'";
                    }
                }

                string name = PrepareName(GetName(parts));
                string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'" + exclude;

                return GetByFilter(filter);
            }
            else
            {
                string name = PrepareName(parts[0].ToString());
                string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'";

                return GetByFilter(filter);
            }
        }
        public Tuple<int, int>[] GetByFilter(string filter, bool range = false)
        {
            if(_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            if (_definedNamesView.Count == 0)
                return null;

            Tuple<int, int>[] o;
            int i = 0;
            if (!range)
            {
                 o = new Tuple<int, int>[_definedNamesView.Count];
                 foreach (DataRowView defName in _definedNamesView)
                     o[i++] = Tuple.Create(int.Parse(defName["R1"].ToString()), int.Parse(defName["C1"].ToString()));
            }
            else
            {
                o = new Tuple<int, int>[2];
                o[0] = Tuple.Create(int.Parse(_definedNamesView[0]["R1"].ToString()), int.Parse(_definedNamesView[0]["C1"].ToString()));
                o[1] = Tuple.Create(int.Parse(_definedNamesView[_definedNamesView.Count - 1]["R1"].ToString()), int.Parse(_definedNamesView[_definedNamesView.Count - 1]["C1"].ToString()));
            }

            return o;
        }
        //public string GetRange(bool excludeDATA0H24, params object[] parts) 
        //{
        //    if (excludeDATA0H24)
        //    {
        //        Array.Resize(ref parts, parts.Length + 2);
        //        parts[parts.Length - 2] = Simboli.EXCLUDE;
        //        parts[parts.Length - 1] = "DATA0.H24";
        //    }
        //    return GetRange(parts);
        //}
        //public string GetRange(params object[] parts)
        //{
        //    int pos = Array.FindIndex(parts, ele => ele.ToString() == Simboli.EXCLUDE);
        //    pos = pos == -1 ? parts.Length : pos;

        //    string exclude = "";
        //    for (int i = pos + 1; i < parts.Length; i++)
        //        exclude += " AND Nome NOT LIKE '%" + parts[i] + "%'";

        //    Array.Resize(ref parts, pos);

        //    string name = PrepareName(GetName(parts));
        //    string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'" + exclude;
        //    Tuple<int, int>[] rng = GetByFilter(filter, true);
        //    return Sheet.R1C1toA1(rng[0].Item1, rng[0].Item2) + ":" + Sheet.R1C1toA1(rng[1].Item1, rng[1].Item2);
        //}

        //public string GetRange(Tuple<int, int> first, Tuple<int, int> last)
        //{
        //    return Sheet.R1C1toA1(first) + ":" + Sheet.R1C1toA1(last);
        //}
        //public string GetRange(Tuple<int,int>[] range)
        //{
        //    return GetRange(range.First(), range.Last());
        //}

        public void ApplySort(string sortCondition)
        {
            _definedNamesView.Sort = sortCondition;
        }

        public bool IsDefined(string name)
        {
            name = PrepareName(name);
            string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'";

            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            return _definedNamesView.Count > 0;
        }
        public bool IsDefined(int row, int column)
        {
            string filter = "Foglio = '" + _foglio + "' AND R1 = " + row + " AND C1 = " + column;
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            return _definedNamesView.Count > 0;
        }

        public List<Tuple<int, int>[]> GetRanges(string name)
        {
            string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'";
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            if (_definedNamesView.Count == 0)
                return null;

            List<Tuple<int, int>[]> o = new List<Tuple<int, int>[]>();

            foreach (DataRowView range in _definedNamesView)
            {
                o.Add(new Tuple<int, int>[2] 
                { 
                    Tuple.Create(int.Parse(_definedNamesView[0]["R1"].ToString()), int.Parse(_definedNamesView[0]["C1"].ToString())),
                    Tuple.Create(int.Parse(_definedNamesView[0]["R2"].ToString()), int.Parse(_definedNamesView[0]["C2"].ToString()))
                });
            }

            return o;
        }

        public bool Editabile(string name)
        {
            name = PrepareName(name);
            string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'";
            if (_definedNamesView.RowFilter != filter)            
                _definedNamesView.RowFilter = filter;
            
            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["Editabile"];
        }
        public bool Editabile(int row, int column)
        {
            string filter = "Foglio = '" + _foglio + "' AND R1 = " + row + " AND C1 = " + column;
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;
            
            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["Editabile"];
        }
        
        public bool SalvaDB(string name)
        {
            name = PrepareName(name);
            string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'";
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;
            
            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["SalvaDB"];
        }
        public bool SalvaDB(int row, int column)
        {
            string filter = "Foglio = '" + _foglio + "' AND R1 = " + row + " AND C1 = " + column;
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["SalvaDB"];
        }
        
        public bool AnnotaModifica(string name)
        {
            name = PrepareName(name);
            string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'";
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["AnnotaModifica"];
        }
        public bool AnnotaModifica(int row, int column)
        {
            string filter = "Foglio = '" + _foglio + "' AND R1 = " + row + " AND C1 = " + column;
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["AnnotaModifica"];
        }

        public DataView GetEditable()
        {
            string filter = "Foglio = '" + _foglio + "' AND Editabile = 1";
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            DataTable dt = new DataTable("Editabili")
            {
                Columns =
                {
                    {"SiglaEntita", typeof(string)},
                    {"SiglaInformazione", typeof(string)},
                    {"SuffissoData", typeof(string)}
                }
            };

            var o =  from DataRow r in _definedNamesView.ToTable("Nome").AsEnumerable().Distinct(new DistinctEntitaInformazione())
                     select dt.LoadDataRow(
                        new object[] 
                            { 
                                r.Field<string>("Nome").Split(Simboli.UNION[0])[0], 
                                r.Field<string>("Nome").Split(Simboli.UNION[0])[1], 
                                r.Field<string>("Nome").Split(Simboli.UNION[0])[2]
                            }, LoadOption.OverwriteChanges);

            return o.CopyToDataTable().DefaultView;
        }

        //public Tuple<int, int>[] GetRangeEntita(object siglaEntita, object suffissoData = null, bool excludeTitleBar = true)
        //{
        //    string filter = "Foglio = '" + _foglio + "' AND Nome LIKE = '" + siglaEntita + "%' AND Nome LIKE '%" + suffissoData + "%'";
        //    if (excludeTitleBar)
        //        filter += " AND Nome NOT LIKE '" + GetName(siglaEntita, "T", suffissoData) + "'";

        //    return GetByFilter(filter);
        //}

        #endregion

        #region Metodi Statici

        /// <summary>
        /// Funzione che prepara il nome per un confronto con l'operatore LIKE. Se il nome passato non fa parte del riepilogo, non è una cella goto, non è un titolo di entita e non finisce con il suffisso data ora, aggiungo un '.' alla fine in maniera da limitare il numero di match.
        /// </summary>
        /// <param name="name">Il nome su cui operare il confronto</param>
        /// <returns>Ritorna la stringa pronta per il confronto con l'operatore LIKE</returns>
        private static string PrepareName(string name)
        {
            //se il nome non fa parte del riepilogo e non finisce con il suffisso data ora, aggiungo un punto
            //if (!Regex.IsMatch(name, @"DATA\d+\.\w+|GRAFICO\d+|RIEPILOGO|DATA\d+\.H\d+|\.T\."))
            if (!Regex.IsMatch(name, @"\.NOTE\.|\.CAMBIO_ASSETTO\.|\.ACCENSIONE\.|GRAFICO\d+|RIEPILOGO|DATA\d+\.H\d+|\.T\.|.+\.GOTO"))
                name += Simboli.UNION;
            return name;
        }
        /// <summary>
        /// Inizializza la tabella dei nomi assegnandole un nome e la restituisce.
        /// </summary>
        /// <param name="name">Il nome da assegnare alla tabella per la serializzazione.</param>
        /// <returns>Ritorna una nuova istanza della tabella dei nomi.</returns>
        public static DataTable GetDefaultTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {Fields.Foglio, typeof(String)},
                        {Fields.Nome, typeof(String)},
                        {Fields.R1, typeof(int)},
                        {Fields.C1, typeof(int)},
                        {Fields.R2, typeof(int)},
                        {Fields.C2, typeof(int)},
                        {Fields.Editabile, typeof(bool)},
                        {Fields.SalvaDB, typeof(bool)},
                        {Fields.AnnotaModifica, typeof(bool)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns[Fields.Foglio], dt.Columns[Fields.Nome] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Funzione che restituisce il nome del foglio a cui appartiene la cella passata in input.
        /// </summary>
        /// <param name="name">Il nome della cella in input.</param>
        /// <returns>Ritorna il nome del foglio a cui appartiene la cella o null se non esiste.</returns>
        public static string GetSheetName(object name)
        {
            DataView definedNamesView = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.NOMI_DEFINITI].DefaultView;
            string filter = "Nome LIKE'" + name + "%'";
            if (definedNamesView.RowFilter != filter)
                definedNamesView.RowFilter = filter;

            if (definedNamesView.Count == 0)
                return null;

            return definedNamesView[0]["Foglio"].ToString();
        }
        /// <summary>
        /// Verifica se il nome in input è definito nella tabella dei nomi per il foglio in input.
        /// </summary>
        /// <param name="sheetName">Il nome del foglio su cui si vuole verificare se la cella è definita</param>
        /// <param name="cellName">Il nome della cella da verificare</param>
        /// <returns>Ritorna true se esiste un match per la coppia foglio - nome, false altrimenti.</returns>
        public static bool IsDefined(string sheetName, string cellName)
        {
            cellName = PrepareName(cellName);
            DataView definedNamesView = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.NOMI_DEFINITI].DefaultView;
            string filter = "Foglio = '" + sheetName + "' AND Nome LIKE '" + cellName + "%'";
            if (definedNamesView.RowFilter != filter)
                definedNamesView.RowFilter = filter;

            return definedNamesView.Count > 0;
        }
        /// <summary>
        /// Verifica se l'indirizzo R-C in input è definito nella tabella dei nomi per il foglio in input.
        /// </summary>
        /// <param name="sheetName">Il nome del foglio su cui si vuole verificare se la cella è definita</param>
        /// <param name="row">La riga dell'indirizzo da verificare</param>
        /// <param name="column">La colonna dell'indirizzo da verificare</param>
        /// <returns>Ritorna true se esiste un match per la coppia foglio - indirizzo, false altrimenti.</returns>
        public static bool IsDefined(string sheetName, int row, int column)
        {
            DataView definedNamesView = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.NOMI_DEFINITI].DefaultView;
            string filter = "Foglio = '" + sheetName + "' AND R1 = " + row + " AND C1 = " + column;
            if (definedNamesView.RowFilter != filter)
                definedNamesView.RowFilter = filter;

            return definedNamesView.Count > 0;
        }
        /// <summary>
        /// Da una lista di oggetti in input, compone il nome con il simbolo di unione.
        /// </summary>
        /// <param name="parts">Lista di stringhe che andranno a comporre il nome in output</param>
        /// <returns>Restituisce la stringa che rappresenta il nome</returns>
        public static string GetName(params object[] parts)
        {
            string o = "";
            bool first = true;
            foreach (object part in parts)
            {
                if (part != null && part.ToString() != "")
                {
                    o += (!first ? Simboli.UNION : "") + part;
                    first = false;
                }
            }
            return o;
        }

        #endregion
    }

    internal class DistinctEntitaInformazione : IEqualityComparer<DataRow>
    {
        public bool Equals(DataRow x, DataRow y)
        {
            string[] xSplit = x["Nome"].ToString().Split(Simboli.UNION[0]);
            string[] ySplit = y["Nome"].ToString().Split(Simboli.UNION[0]);

            return xSplit[0] == ySplit[0] && xSplit[1] == ySplit[1] && xSplit[2] == ySplit[2];
        }

        public int GetHashCode(DataRow obj)
        {
            string[] objSplit = obj["Nome"].ToString().Split(Simboli.UNION[0]);
            string o = DefinedNames.GetName(objSplit[0], objSplit[1], objSplit[2]);
            return o.GetHashCode();
        }
    }




}
