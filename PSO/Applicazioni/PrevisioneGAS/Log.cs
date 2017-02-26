using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    public partial class Log
    {
        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(this.Log_Startup);
        }

        #endregion

        #region Callbacks

        private void Log_Startup(object sender, EventArgs e)
        {
            Unprotect(PSO.Base.Workbook.Password);
            this.logList.DataSource = PSO.Base.Workbook.LogDataTable;

            this.logList.AutoSetDataBoundColumnHeaders = true;
            this.logList.Range.EntireColumn.AutoFit();
            this.logList.TableStyle = "TableStyleLight16";

            ((Excel.Range)Columns[2]).NumberFormat = "dd/MM/yyyy";
            ((Excel.Range)Columns[2]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            Protect(PSO.Base.Workbook.Password, allowSorting: true, allowFiltering: true);
        }

        #endregion

    }
}
