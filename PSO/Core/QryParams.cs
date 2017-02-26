using System.Collections;
using System.Collections.Generic;

namespace Iren.PSO.Core
{
    public class QryParams : IEnumerable
    {
        #region Variabili

        Dictionary<string, object> _parameters = new Dictionary<string, object>();

        #endregion

        #region Costruttori

        public QryParams() {}

        #endregion

        #region Proprietà

        public object this[string key]
        {
            get
            {
                return _parameters[key];
            }
            set
            {
                _parameters[key] = value;
            }
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Aggiunge un nuovo parametro con chiave key e valore value.
        /// </summary>
        /// <param name="key">Chiave.</param>
        /// <param name="value">Valore.</param>
        public void Add(string key, object value)
        {
            _parameters.Add(key, value);
        }
        /// <summary>
        /// Verifica se contiente la chiave key.
        /// </summary>
        /// <param name="key">Chiave.</param>
        /// <returns>True se la chiave esiste, false altrimenti.</returns>
        public bool ContainsKey(string key)
        {
            return _parameters.ContainsKey(key);
        }

        public IEnumerator GetEnumerator() { return _parameters.GetEnumerator(); }

        public override string ToString()
        {
            if (_parameters.Count == 0)
                return "params:{}";

            string o = "";
            foreach (var kv in _parameters)
            {
                o += kv.Key + "=" + kv.Value + ",";
            }
            return "params:{" + o.Substring(0, o.Length - 1) + "}";
        }

        #endregion
    }
}
