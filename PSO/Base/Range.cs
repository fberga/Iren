using System;
using System.Collections;
using System.Text.RegularExpressions;

namespace Iren.PSO.Base
{
    public class Range
    {
        #region Variabili

        private int _startRow;
        private int _startColumn;
        private int _rowOffset = 1;
        private int _colOffset = 1;

        private RowsCollection _rows;
        private ColumnsCollection _cols;
        private CellsCollection _cells;

        #endregion

        #region Proprietà

        public int StartRow
        {
            get { return _startRow; }
            set { _startRow = value; }
        }
        public int StartColumn
        {
            get { return _startColumn; }
            set { _startColumn = value; }
        }
        public int RowOffset
        {
            get { return _rowOffset; }
            set { _rowOffset = value < 1 ? 1 : value; }
        }
        public int ColOffset
        {
            get { return _colOffset; }
            set { _colOffset = value < 1 ? 1 : value; }
        }

        public int EndRow
        {
            get { return _startRow + _rowOffset; }
        }
        public int EndColumn
        {
            get { return _startColumn + _colOffset; }
        }

        public RowsCollection Rows
        {
            get { return _rows; }
        }
        public ColumnsCollection Columns
        {
            get { return _cols; }
        }
        public CellsCollection Cells
        {
            get { return _cells; }
        }

        public bool Lock { get; set; }

        #endregion

        #region Costruttori

        /// <summary>
        /// Costruttore vuoto. Crea l'oggetto inizializzando a vuote le collezioni di righe, colonne e celle.
        /// </summary>
        public Range() 
        {
            _rows = new RowsCollection(this);
            _cols = new ColumnsCollection(this);
            _cells = new CellsCollection(this);
        }
        /// <summary>
        /// Costruttore di copia. Crea l'oggetto copiando le informazioni da oth.
        /// </summary>
        /// <param name="oth">Oggetto da cui copiare il range.</param>
        public Range(Range oth) 
            : this()
        {
            _startRow = oth.StartRow;
            _startColumn = oth.StartColumn;
            _rowOffset = oth.RowOffset;
            _colOffset = oth.ColOffset;
        }
        /// <summary>
        /// Definisce un range cella situato alle coordinate (row, column).
        /// </summary>
        /// <param name="row">Riga.</param>
        /// <param name="column">Colonna.</param>
        public Range(int row, int column)
            : this()
        {
            _startRow = row;
            _startColumn = column;
        }
        /// <summary>
        /// Definisce un range di una colonna che parte da (row, column) e si estende per rowOffset righe. Quando rowOffset è a 1 il range è una cella.
        /// </summary>
        /// <param name="row">Riga.</param>
        /// <param name="column">Colonna.</param>
        /// <param name="rowOffset">Righe di offset. Se 1 inizializzo un range cella.</param>
        public Range(int row, int column, int rowOffset)
            : this()
        {
            _startRow = row;
            _startColumn = column;
            _rowOffset = rowOffset;
        }
        /// <summary>
        /// Definisce un range generico che parte da (row, column) e si estende per rowOffset righe e colOffset colonne. Se rowOffset e colOffset sono a 1 il range è di una cella.
        /// </summary>
        /// <param name="row">Riga.</param>
        /// <param name="column">Colonna.</param>
        /// <param name="rowOffset">Righe di offset.</param>
        /// <param name="colOffset">Colonne di offset.</param>
        public Range(int row, int column, int rowOffset, int colOffset)
            : this()
        {
            _startRow = row;
            _startColumn = column;
            _rowOffset = rowOffset;
            _colOffset = colOffset;
        }
        /// <summary>
        /// Definisce un range a partire da una stringa di indirizzo in forma A1.
        /// </summary>
        /// <param name="range">Indirizzo in forma A1.</param>
        public Range(string range)
            : this()
        {
            Range rng = A1toRange(range);

            _startRow = rng.StartRow;
            _startColumn = rng.StartColumn;
            _rowOffset = rng.RowOffset;
            _colOffset = rng.ColOffset;
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Prta il range a rowOffset e colOffset. Restituisce il nuovo oggetto range esteso.
        /// </summary>
        /// <param name="rowOffset">Nuovo offset di righe.</param>
        /// <param name="colOffset">Nuovo offset di colonne.</param>
        /// <returns>Il nuovo oggetto range esteso.</returns>
        public Range Extend(int rowOffset = 1, int colOffset = 1) 
        {
            RowOffset = rowOffset;
            ColOffset = colOffset;
            
            return this;
        }
        /// <summary>
        /// Estende il range di rowOffset e colOffset. Restituisce il nuovo oggetto range esteso.
        /// </summary>
        /// <param name="rowOffset">Righe di offset da aggiungere.</param>
        /// <param name="colOffset">Colonne di offset da aggiungere.</param>
        /// <returns>Il nuovo oggetto range esteso.</returns>
        public Range ExtendOf(int rowOffset = 0, int colOffset = 0)
        {
            _rowOffset += rowOffset;
            _colOffset += colOffset;

            return this;
        }
        /// <summary>
        /// Scrive il range in formato A1.
        /// </summary>
        /// <returns>La stringa di indirizzo in formato A1.</returns>
        public override string ToString()
        {
            return GetRange(_startRow, _startColumn, _rowOffset, _colOffset, Lock);
        }
        /// <summary>
        /// Verifica se il range contiene rng.
        /// </summary>
        /// <param name="rng">Range di cui verificare l'appartenenza.</param>
        /// <returns>True se rng è un sottorange al massimo uguale, false altrimenti</returns>
        public bool Contains(Range rng)
        {
            return StartRow <= rng.StartRow
                && StartColumn <= rng.StartColumn 
                && StartRow + RowOffset >= rng.StartRow + rng.RowOffset 
                && StartColumn + ColOffset >= rng.StartColumn + rng.ColOffset;
        }

        #endregion

        #region Metodi Statici

        /// <summary>
        /// Converte l'indirizzo dal formato RC al formato A1.
        /// </summary>
        /// <param name="riga">Riga.</param>
        /// <param name="colonna">Colonna.</param>
        /// <returns>La stringa che rappresenta l'indirizzo in formato A1.</returns>
        public static string R1C1toA1(int riga, int colonna, bool lk = false)
        {
            string output = "";
            while (colonna > 0)
            {
                int lettera = (colonna - 1) % 26;
                output = Convert.ToChar(lettera + 65) + output;
                colonna = (colonna - lettera) / 26;
            }
            output += (lk ? "$" : "") + riga;
            return (lk ? "$" : "") + output;
        }
        /// <summary>
        /// Converte l'indirizzo in formato A1 in un oggetto Range.
        /// </summary>
        /// <param name="address">Indirizzo in formato A1.</param>
        /// <returns>Oggetto range corrispondente all'indirizzo.</returns>
        public static Range A1toRange(string address)
        {
            string[] parts = address.Split(':');

            int[] rows = new int[parts.Length];
            int[] cols = new int[parts.Length];
            int j = 0;
            foreach (string part in parts)
            {
                string tmp = part.Replace("$", "");
                string alpha = Regex.Match(tmp, @"\D+").Value;
                rows[j] = int.Parse(Regex.Match(tmp, @"\d+").Value);

                cols[j] = 0;
                int incremento = (alpha.Length == 1 ? 1 : 26 * (alpha.Length - 1));
                for (int i = 0; i < alpha.Length; i++)
                {
                    cols[j] += (char.ConvertToUtf32(alpha, i) - 64) * incremento;
                    incremento = incremento - 26 == 0 ? 1 : incremento - 26;
                }
                j++;
            }

            Range rng = new Range();
            rng.StartRow = rows[0];
            rng.StartColumn = cols[0];
            
            if (rows.Length == 2)
            {
                rng.RowOffset = rows[1] - rows[0] + 1;
                rng.ColOffset = cols[1] - cols[0] + 1;
            }

            return rng;
        }
        /// <summary>
        /// Crea l'indirizzo in formato A1 a partire da tutti i dati che compongono il range.
        /// </summary>
        /// <param name="row">Riga.</param>
        /// <param name="column">Colonna.</param>
        /// <param name="rowOffset">Righe di offset.</param>
        /// <param name="colOffset">Colonne di offset.</param>
        /// <returns>Stringa di indirizzo in formato A1.</returns>
        public static string GetRange(int row, int column, int rowOffset = 1, int colOffset = 1, bool lk = false)
        {
            if ((rowOffset == 1 && colOffset == 1))
                return R1C1toA1(row, column, lk);

            return R1C1toA1(row, column, lk) + ":" + R1C1toA1(row + rowOffset - 1, column + colOffset - 1, lk);
        }

        #endregion

        #region Classi Interne

        public class RowsCollection : IEnumerable
        {
            private Range _r;

            internal RowsCollection(Range r)
            {
                _r = r;
            }

            public Range this[int row]
            {
                get
                {
                    return new Range(_r.StartRow + row, _r.StartColumn, 1, _r.ColOffset);
                }
            }
            public Range this[int row1, int row2]
            {
                get
                {
                    return new Range(_r.StartRow + row1, _r.StartColumn, row2 - row1 + 1, _r.ColOffset);
                }
            }
            public int Count
            {
                get
                {
                    return _r.RowOffset;
                }
            }
            public IEnumerator GetEnumerator()
            {
                return new RowsEnum(_r);
            }
        }
        public class ColumnsCollection : IEnumerable
        {
            private Range _r;

            internal ColumnsCollection(Range r)
            {
                _r = r;
            }

            public Range this[int column]
            {
                get
                {
                    return new Range(_r.StartRow, _r.StartColumn + column, _r.RowOffset, 1);
                }
            }
            public Range this[int col1, int col2]
            {
                get
                {
                    return new Range(_r.StartRow, _r.StartColumn + col1, _r.RowOffset, col2 - col1 + 1);
                }
            }
            public int Count
            {
                get
                {
                    return _r.ColOffset;
                }
            }
            public IEnumerator GetEnumerator()
            {
                return new ColumnsEnum(_r);
            }
        }
        public class CellsCollection
        {
            private Range _r;

            internal CellsCollection(Range r)
            {
                _r = r;
            }

            public Range this[int row, int column]
            {
                get
                {
                    return new Range(_r.StartRow + row, _r.StartColumn + column);
                }
            }
            public int Count
            {
                get
                {
                    return _r.ColOffset * _r.RowOffset;
                }
            }
            public IEnumerator GetEnumerator()
            {
                return new CellsEnum(_r);
            }
        }

        public class RowsEnum : IEnumerator
        {
            Range _r;
            int _position = -1;
            int _maxOffset = -1;

            public RowsEnum(Range r)
            {
                _r = r;
                _maxOffset = _r.RowOffset;
                
            }

            public object Current
            {
                get { return _r.Rows[_position]; }
            }

            public bool MoveNext()
            {
                _position++;
                return _position < _maxOffset;
            }

            public void Reset()
            {
                _position = -1;
            }
        }
        public class ColumnsEnum : IEnumerator
        {
            Range _r;
            int _position = -1;
            int _maxOffset = -1;

            public ColumnsEnum(Range r)
            {
                _r = r;
                _maxOffset = _r.ColOffset;

            }

            public object Current
            {
                get { return _r.Columns[_position]; }
            }

            public bool MoveNext()
            {
                _position++;
                return _position < _maxOffset;
            }

            public void Reset()
            {
                _position = -1;
            }
        }
        public class CellsEnum : IEnumerator
        {
            Range _r;
            int _xPosition = -1;
            int _yPosition = 0;
            int _xOffset = -1;
            int _yOffset = -1;

            public CellsEnum(Range r)
            {
                _r = r;
                _xOffset = _r.ColOffset;
                _yOffset = _r.RowOffset;

            }

            public object Current
            {
                get { return _r.Cells[_yPosition, _xPosition]; }
            }

            public bool MoveNext()
            {
                _xPosition++;
                if (_xPosition == _xOffset)
                {
                    _xPosition = 0;
                    _yPosition++;
                }
                return _yPosition < _yOffset;
            }

            public void Reset()
            {
                _xPosition = -1;
                _yPosition = 0;
            }
        }
        #endregion
    }
}
