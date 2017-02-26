using System;
using System.Data;
using System.Data.SqlClient;

namespace Iren.PSO.Core
{
    class Command : IDisposable
    {
        #region Variabili

        private SqlConnection _sqlConn;

        #endregion

        #region Costruttori

        public Command(SqlConnection sqlConn) 
        {
            _sqlConn = sqlConn;
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Restituisce il comando SQL. Non accetta parametri in caso di esecuzione di funzioni/stored procedure.
        /// </summary>
        /// <param name="commandText">Testo del comando. Può essere SQL o nome funzione/stored procedure</param>
        /// <param name="commandType">Tipo di comando. Deve corrispondere a quello che si trova nel testo del comando.</param>
        /// <param name="timeout">Timeout di esecuzione.</param>
        /// <returns>Comando SqlCommand.</returns>
        public SqlCommand SqlCmd(string commandText, CommandType commandType, int timeout = 300)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = _sqlConn;
            cmd.CommandText = commandText;
            cmd.CommandType = commandType;
            cmd.CommandTimeout = timeout;
            return cmd;
        }
        /// <summary>
        /// Restituisce il comando SQL presupponendo si tratti di una stored procedure. Non accetta parametri.
        /// </summary>
        /// <param name="commandText">Testo del comando (deve essere una stored procedure).</param>
        /// <param name="timeout">Timeout di esecuzione.</param>
        /// <returns>Comando SqlCommand.</returns>
        public SqlCommand SqlCmd(string commandText, int timeout = 300)
        {
            return SqlCmd(commandText, CommandType.StoredProcedure, timeout);
        }
        /// <summary>
        /// Restituisce il comando SQL. Accetta una lista di parametri.
        /// </summary>
        /// <param name="commandText">Testo del comando.</param>
        /// <param name="commandType">Tipo di comando.</param>
        /// <param name="parameters">Lista di parametri.</param>
        /// <param name="timeout">Timeout di esecuzione.</param>
        /// <returns>Comando SqlCommand.</returns>
        public SqlCommand SqlCmd(string commandText, CommandType commandType, QryParams parameters, int timeout = 300)
        {
            SqlCommand cmd = SqlCmd(commandText, commandType, timeout);
            try
            {
                SqlCommandBuilder.DeriveParameters(cmd);
                foreach (SqlParameter par in cmd.Parameters)
                {
                    if(parameters.ContainsKey(par.ParameterName))
                        par.Value = parameters[par.ParameterName];                    
                }
            }
            catch (Exception)
            {                
            }
            return cmd;
        }
        /// <summary>
        /// Restituisce il comando SQL presupponendo si tratti di una stored procedure. Accetta una lista di parametri.
        /// </summary>
        /// <param name="commandText">Testo del comando.</param>
        /// <param name="parameters">Lista di parametri.</param>
        /// <param name="timeout">Timeout di esecuzione.</param>
        /// <returns>Comando SqlCommand.</returns>
        public SqlCommand SqlCmd(string commandText, QryParams parameters, int timeout = 300)
        {
            return SqlCmd(commandText, CommandType.StoredProcedure, parameters, timeout);
        }
       
        #endregion


        public void Dispose()
        {
            _sqlConn.Dispose();
        }
    }
}
