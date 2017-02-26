using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Base
{
    public interface IPSOThisWorkbook
    {
        System.Version Version { get; }
        Microsoft.Office.Tools.Excel.Workbook Base { get; }
        Microsoft.Office.Interop.Excel.Worksheet ActiveSheet { get; }        
        Excel.Sheets Sheets { get; }
        Excel.Application Application { get; }

        string Name { get; }
        string Path { get; }
        string FullName { get; }
        string Pwd { get; }
        string NomeUtente { get; set; }
        string Ambiente { get; set; }

        int IdApplicazione { get; set; }
        int IdUtente { get; set; }
        int IdStagione { get; set; }

        DateTime DataAttiva { get; set; }

        System.Data.DataSet RepositoryDataSet { get; }
        System.Data.DataTable LogDataTable { get; set; }
        System.Data.DataSet RibbonDataSet { get; }
    }
}
