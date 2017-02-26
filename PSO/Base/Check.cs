using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Base
{
    /// <summary>
    /// Classe base per la gestione dei check. Fornisce il launcher che viene richiamato nel sistema e le funzioni di base.
    /// </summary>
    public class Check
    {
        #region Variabili
        /// <summary>
        /// Foglio di lavoro a cui appartengono le funzioni di check.
        /// </summary>
        protected Excel.Worksheet _ws;
        /// <summary>
        /// Struttura di indicizzazione dei nomi.
        /// </summary>
        protected DefinedNames _nomiDefiniti;
        /// <summary>
        /// Oggetto di check.
        /// </summary>
        protected CheckObj _check;

        #endregion

        #region Metodi
        /// <summary>
        /// Implementazione di base del launcher. Restituisce un oggetto CheckOutput vuoto. Il CheckOutput contiene il nodo della TreeView che ha le informazioni del check e lo stato in cui si trova. La classe va sovrascritta in ogni applicativo che contenga funzioni di check.
        /// </summary>
        /// <param name="ws">Foglio di lavoro su cui calcolare i check.</param>
        /// <param name="definedNames">Struttura di indicizzazione dei nomi per il foglio.</param>
        /// <param name="check">Oggetto di check del quale eseguire il controllo.</param>
        /// <returns>Oggetto CheckOutput vuoto.</returns>
        public virtual CheckOutput ExecuteCheck(Excel.Worksheet ws, DefinedNames definedNames, CheckObj check)
        {
            return new CheckOutput();
        }
        
        /// <summary>
        /// Utilizzando l'indicizzazione restituisce il valore della cella convertito in Decimal.
        /// </summary>
        /// <param name="siglaEntita">La sigla dell'entita per l'indicizzazione.</param>
        /// <param name="siglaInformazione">La sigla dell'informazione per l'indicizzazione.</param>
        /// <param name="suffissoData">Il suffisso data per l'indicizzazione.</param>
        /// <param name="suffissoOra">Il suffisso ora per l'indicizzazione.</param>
        /// <returns>Il valore della cella convertito in Decimal</returns>
        protected virtual decimal GetDecimal(object siglaEntita, object siglaInformazione, object suffissoData, object suffissoOra)
        {
            object tmp = _ws.Range[_nomiDefiniti.Get(siglaEntita, siglaInformazione, suffissoData, suffissoOra).ToString()].Value;

            if (tmp == null || tmp.Equals(""))
                return (decimal)0;

            return Convert.ToDecimal(tmp);
        }
        /// <summary>
        /// Utilizzando un range restituisce il valore della cella convertito in Decimal.
        /// </summary>
        /// <param name="rng">Range di cui estrarre il valore.</param>
        /// <returns>Il valore della cella convertito in Decimal.</returns>
        protected virtual decimal GetDecimal(Range rng)
        {
            object tmp = _ws.Range[rng.ToString()].Value;

            if (tmp == null || tmp.Equals(""))
                return (decimal)0;

            return Convert.ToDecimal(tmp);
        }
        /// <summary>
        /// Utilizzando l'indicizzazione restituisce il valore della cella.
        /// </summary>
        /// <param name="siglaEntita">La sigla dell'entita per l'indicizzazione.</param>
        /// <param name="siglaInformazione">La sigla dell'informazione per l'indicizzazione.</param>
        /// <param name="suffissoData">Il suffisso data per l'indicizzazione.</param>
        /// <param name="suffissoOra">Il suffisso ora per l'indicizzazione.</param>
        /// <returns>Il valore della cella.</returns>
        protected virtual object GetObject(object siglaEntita, object siglaInformazione, object suffissoData, object suffissoOra)
        {
            return _ws.Range[_nomiDefiniti.Get(siglaEntita, siglaInformazione, suffissoData, suffissoOra).ToString()].Value;
        }
        /// <summary>
        /// Utilizzando un range restituisce il valore della cella.
        /// </summary>
        /// <param name="rng">Range di cui estrarre il valore.</param>
        /// <returns>Il valore della cella.</returns>
        protected virtual object GetObject(Range rng)
        {
            return _ws.Range[rng.ToString()].Value;
        }
        /// <summary>
        /// Utilizzando l'indicizzazione restituisce il valore della cella convertito in String.
        /// </summary>
        /// <param name="siglaEntita">La sigla dell'entita per l'indicizzazione.</param>
        /// <param name="siglaInformazione">La sigla dell'informazione per l'indicizzazione.</param>
        /// <param name="suffissoData">Il suffisso data per l'indicizzazione.</param>
        /// <param name="suffissoOra">Il suffisso ora per l'indicizzazione.</param>
        /// <returns>Il valore della cella convertito in String</returns>
        protected virtual string GetString(object siglaEntita, object siglaInformazione, object suffissoData, object suffissoOra)
        {
            return (string)(_ws.Range[_nomiDefiniti.Get(siglaEntita, siglaInformazione, suffissoData, suffissoOra).ToString()].Value ?? "");
        }
        /// <summary>
        /// Utilizzando un range restituisce il valore della cella convertito in String.
        /// </summary>
        /// <param name="rng">Range di cui estrarre il valore.</param>
        /// <returns>Il valore della cella convertito in String.</returns>
        protected virtual string GetString(Range rng)
        {
            return (string)(_ws.Range[rng.ToString()].Value ?? "");
        }
        /// <summary>
        /// Definisce la formattazione degli elementi in errore nella TreeView.
        /// </summary>
        /// <param name="node">Il nodo da formattare della TreeView</param>
        protected virtual void ErrorStyle(ref TreeNode node)
        {
            node.BackColor = System.Drawing.Color.Red;
            node.ForeColor = System.Drawing.Color.Yellow;
            node.NodeFont = new System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Bold);
        }
        /// <summary>
        /// Definisce la formattazione degli elementi in attenzione nella TreeView.
        /// </summary>
        /// <param name="node">Il nodo da formattare della TreeView</param>
        protected virtual void AlertStyle(ref TreeNode node)
        {
            node.BackColor = System.Drawing.Color.Yellow;
            node.ForeColor = System.Drawing.Color.Red;
            node.NodeFont = new System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Bold);
        }

        #endregion
    }

    public class CheckOutput
    {
        /// <summary>
        /// Gli stati del check.
        /// </summary>
        public enum CheckStatus
        {
            Ok, Alert, Error
        }

        #region Variabili

        /// <summary>
        /// Il nodo della TreeView collegato al check.
        /// </summary>
        TreeNode _node;
        /// <summary>
        /// Lo stato del Check.
        /// </summary>
        CheckStatus _status;

        #endregion

        #region Costruttori

        public CheckOutput()
        {
            _node = new TreeNode();
            _status = CheckStatus.Ok;
        }

        public CheckOutput(TreeNode node, CheckStatus status)
        {
            _node = node;
            _status = status;
        }

        #endregion

        #region Proprietà

        /// <summary>
        /// Restituisce il nodo della TreeView collegato al check.
        /// </summary>
        public TreeNode Node { get { return _node; } }
        /// <summary>
        /// Restituisce lo stato del Check del tipo CheckOutput.CheckStatus.Ok, CheckOutput.CheckStatus.Alert, CheckOutput.CheckStatus.Error.
        /// </summary>
        public CheckStatus Status { get { return _status; } }

        #endregion
    }

    public class CheckObj
    {
        #region Proprietà

        public string SiglaEntita { get; set; }
        public string Range { get; set; }
        public int Type { get; set; }

        #endregion

        #region Costruttori

        public CheckObj(string range)
        {
            Range = range;
        }
        public CheckObj(string siglaEntita, string range, int type)
        {
            SiglaEntita = siglaEntita;
            Range = range;
            Type = type;
        }

        #endregion
    }
}

