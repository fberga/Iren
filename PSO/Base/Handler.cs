using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Base
{
    /// <summary>
    /// Classe che gestisce molti dei comportamenti standard del workbook.
    /// </summary>
    public class Handler
    {
        #region Metodi 

        /// <summary>
        /// Handler per il click su celle di selezione.
        /// </summary>
        /// <param name="Sh"></param>
        /// <param name="Target"></param>
        public static void SelectionClick(object Sh, Excel.Range Target)
        {
            DefinedNames definedNames = new DefinedNames(Target.Worksheet.Name, DefinedNames.InitType.Selection);
            Range rng = new Range(Target.Row, Target.Column);
            Selection sel;
            int val;
            if (definedNames.HasSelections() && definedNames.TryGetSelectionByPeer(rng, out sel, out val))
            {
                Target.Worksheet.Unprotect(Workbook.Password);
                if (sel != null)
                {
                    Workbook.ScreenUpdating = false;

                    //evito di annotare il cambiamento della cella di selezione: non ha senso e cmq va in errore perché il simbolo non viene riconosciuto
                    if (Simboli.ModificaDati)
                        //Workbook.WB.SheetChange -= StoreEdit;
                        Workbook.RemoveStdStoreEdit();


                    sel.ClearSelections(Target.Worksheet);
                    sel.Select(Target.Worksheet, rng.ToString());

                    //annoto modifiche e le salvo sul DB
                    //Workbook.WB.SheetChange += StoreEdit;
                    Workbook.AddStdStoreEdit();

                    Target.Worksheet.Range[sel.RifAddress].Value = val;
                    
                    //se non ero in modifica tolgo l'handler alla modifica delle celle
                    if (!Simboli.ModificaDati)
                        //Workbook.WB.SheetChange -= StoreEdit;
                        Workbook.RemoveStdStoreEdit();

                    DataBase.SalvaModificheDB();
                    Workbook.ScreenUpdating = true;
                }
                Target.Worksheet.Protect(Workbook.Password);
            }
        }
        /// <summary>
        /// Gestisce il caso in cui ci sia una selezione multipla che andrebbe a scrivere su righe nascoste: allerta l'utente e impedisce di procedere con la modifica.
        /// </summary>
        /// <param name="Sh">Sheet di provenienza.</param>
        /// <param name="Target">Range selezionato dall'utente.</param>
        public static void CellClick(object Sh, Excel.Range Target)
        {
            //controllo che la selezione non sia multi-linea con in mezzo delle righe nascoste - nel caso avverto l'utente che non può effettuare modifiche
            if (Target.Rows.Count > 1)
            {
                if(Simboli.ModificaDati)
                {
                    foreach (Excel.Range r in Target.Rows)
                    {
                        if (r.EntireRow.Hidden)
                        {
                            System.Windows.Forms.MessageBox.Show("Nella selezione sono incluse righe nascoste. Non si può procedere con la modifica...", Simboli.NomeApplicazione + " - ATTENZIONE", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);

                            Target.Cells[1, 1].Select();

                            break;
                        }
                    }
                }
            }
            else
            {
                try
                {
                    DefinedNames newDefinedNames = new DefinedNames(Target.Worksheet.Name, DefinedNames.InitType.GOTOs);
                    string address = newDefinedNames.GetGotoFromAddress(Range.R1C1toA1(Target.Row, Target.Column));
                    Goto(address);
                }
                catch {}
            }
        }
        /// <summary>
        /// Sposta la selezione su address e la centra nello schermo.
        /// </summary>
        /// <param name="address">L'indirizzo della cella/range da selezionare in forma A1</param>
        public static void Goto(string address)
        {
            if (address != "")
            {
                Excel.Range rng = (Excel.Range)Workbook.Application.Range[address];
                ((Excel._Worksheet)rng.Worksheet).Activate();
                rng.Select();
                Workbook.Application.ActiveWindow.SmallScroll(rng.Row - Workbook.Application.ActiveWindow.VisibleRange.Cells[1, 1].Row - 1);
            }
        }
        /// <summary>
        /// Funzione per il salvataggio delle modifiche apportate a ranges anche non contigui.
        /// </summary>
        /// <param name="Target">L'insieme dei ranges modificati</param>
        /// <param name="annotaModifica">Se la modifica va segnalata all'utente attraverso il commento sulla cella oppure no.</param>
        /// <param name="fromCalcolo">Flag per eseguire azioni particolari nel caso la provenienza del salvataggio sia da un calcolo.</param>
        /// <param name="tableName">La tabella in cui inserire le modifiche. Di default Tab.Modifica. Utile specificarme una diversa nel caso di esportazione XML.</param>
        public static void StoreEdit(Excel.Range Target, int annotaModifica = -1, bool fromCalcolo = false, string tableName = DataBase.TAB.MODIFICA)
        {
            if (Workbook.IdUtente != 0 && Workbook.CategorySheets.Contains(Target.Worksheet))        //non salva sulla tabella delle modifiche se l'utente non è configurato
            {
                Excel.Worksheet ws = Target.Worksheet;
                bool wasProtected = ws.ProtectContents;
                bool screenUpdating = ws.Application.ScreenUpdating;
                if (wasProtected)
                    ws.Unprotect(Workbook.Password);

                if (screenUpdating)
                    Workbook.ScreenUpdating = false;

                DefinedNames definedNames = new DefinedNames(Target.Worksheet.Name, DefinedNames.InitType.SaveDB);
                DataTable dt = Workbook.Repository[tableName];

                if (ws.ChartObjects().Count > 0 && !fromCalcolo)
                {
                    Sheet s = new Sheet(ws);
                    s.AggiornaGrafici();
                }

                string[] ranges = Target.Address.Split(',');
                foreach (string range in ranges)
                {
                    Range rng = new Range(range);
                    Range merged = null;
                    object mergedVal = null;
                    try
                    {   //controllo se c'è un merge nel range
                        merged = new Range(ws.Range[rng.ToString()].MergeArea.Address);
                        //salvo il valore
                        mergedVal = ws.Range[rng.ToString()].Value;
                        rng = merged;
                    }
                    catch { }

                    
                    foreach (Range row in rng.Rows)
                    {
                        if (definedNames.SaveDB(row.StartRow))
                        {
                            bool annota = annotaModifica == -1 ? definedNames.ToNote(row.StartRow) : annotaModifica == 1;
                            foreach (Range column in row.Columns)
                            {
                                string[] parts = definedNames.GetNameByAddress(column.StartRow, column.StartColumn).Split(Simboli.UNION[0]);

                                string data;
                                if (parts.Length == 4)
                                    data = Date.GetDataFromSuffisso(parts[2], parts[3]);
                                else
                                    data = Date.GetDataFromSuffisso(parts[2], "");

                                if (!Workbook.Application.WorksheetFunction.IsErr(ws.Range[column.ToString()]))
                                {
                                    DataRow r = dt.Rows.Find(new object[] { parts[0], parts[1], data });
                                    if (r != null) 
                                    {
                                        object val = ws.Range[column.ToString()].Value ?? "";

                                        if (merged == null)
                                            r["Valore"] = (val.Equals("-") ? "0" : val);
                                        else
                                            r["Valore"] = mergedVal ?? "";
                                    }
                                    else
                                    {
                                        DataRow newRow = dt.NewRow();

                                        object val = ws.Range[column.ToString()].Value ?? "";

                                        newRow["SiglaEntita"] = parts[0];
                                        newRow["SiglaInformazione"] = parts[1];
                                        newRow["Data"] = data;
                                        if (merged == null)
                                            newRow["Valore"] = (val.Equals("-") ? "0" : val);
                                        else
                                            newRow["Valore"] = mergedVal ?? "";
                                        newRow["AnnotaModifica"] = annota ? "1" : "0";
                                        newRow["IdApplicazione"] = Workbook.IdApplicazione;
                                        newRow["IdUtente"] = Workbook.IdUtente;

                                        dt.Rows.Add(newRow);
                                    }

                                    if (annota)
                                    {
                                        ws.Range[column.ToString()].ClearComments();
                                        ws.Range[column.ToString()].AddComment("Valore inserito manualmente").Visible = false;
                                    }
                                }
                            }
                        }
                    }
                }

                if (wasProtected)
                    ws.Protect(Workbook.Password);

                if (screenUpdating)
                    ws.Application.ScreenUpdating = true;
            }
        }
        /// <summary>
        /// Funzione per il salvataggio delle modifiche apportate dall'utente quando la modifica è abilitata.
        /// </summary>
        /// <param name="Sh">Sheet.</param>
        /// <param name="Target">Range.</param>
        public static void StoreEdit(object Sh, Excel.Range Target)
        {
            StoreEdit(Target);
        }
        
        /// <summary>
        /// Funzione per il salvataggio dei valori originali del foglio prima di aver salvato le modifiche di incremento.
        /// </summary>
        /// <param name="Target">Sheet.</param>
        /// <param name="tableName">Range.</param>
        /// <param name="Categoria">Range.</param>
        public static void SaveOriginValues(Excel.Range Target, string tableName, string Categoria = "")
        {
            DefinedNames definedNames = new DefinedNames(Target.Worksheet.Name, DefinedNames.InitType.SaveDB);
           
            DataTable dt = Workbook.Repository[tableName];

            foreach (Excel.Range r in Target)
            {
                string[] parts = definedNames.GetNameByAddress(r.Row, r.Column).Split(Simboli.UNION[0]);

                string data;
                if (parts.Length == 4)
                    data = Date.GetDataFromSuffisso(parts[2], parts[3]);
                else
                    data = Date.GetDataFromSuffisso(parts[2], "");

                if (dt.Rows.Find(new object[] { parts[0], parts[1], data }) == null)
                {
                    dt.Rows.Add(string.IsNullOrEmpty(Categoria) ? "" : Categoria, parts[0], parts[1], data, r.Value2, r.Comment == null ? "" : r.Comment.Text());
                }
            }
        }

        /// <summary>
        /// Handler per cambiare il label di modifica e la scritta sotto il tasto sul ribbon.
        /// </summary>
        /// <param name="modifica">True se modifica è abilitata, false se disabilitata.</param>
        public static void ChangeModificaDati(bool modifica)
        {
            Excel.Worksheet ws = Workbook.Sheets["Main"];
            ws.Shapes.Item("lbModifica").Locked = false;
            ws.Shapes.Item("lbModifica").TextFrame.Characters().Text = "Modifica dati: " + (modifica ? "SI" : "NO");
            if (modifica) 
            {
                //giallo
                ws.Shapes.Item("lbModifica").Line.Weight = 2f;
                ws.Shapes.Item("lbModifica").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 204, 0));
                ws.Shapes.Item("lbModifica").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 102));
            }
            else
            {
                //bianco normale
                ws.Shapes.Item("lbModifica").Line.Weight = 0.75f;
                ws.Shapes.Item("lbModifica").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                ws.Shapes.Item("lbModifica").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 0, 0));
                ws.Shapes.Item("lbModifica").Line.ForeColor.Brightness = +0.75f;
            }
            ws.Shapes.Item("lbModifica").Locked = true;
        }
        /// <summary>
        /// Handler per cambiare il label dell'ambiente.
        /// </summary>
        /// <param name="ambiente">Sigla Ambiente.</param>
        public static void ChangeAmbiente(string ambiente)
        {
            bool isProt = Sheet.Protected;
            if (isProt)
                Sheet.Protected = false;

            Workbook.Main.Shapes.Item("lbTest").Locked = false;
            switch (ambiente)
            {
                case Simboli.DEV:
                    Workbook.Main.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: DEVELOPMENT";
                    //rosso
                    Workbook.Main.Shapes.Item("lbTest").Line.Weight = 2f;
                    Workbook.Main.Shapes.Item("lbTest").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(140, 56, 54));
                    Workbook.Main.Shapes.Item("lbTest").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 80, 77));
                    break;
                case Simboli.TEST:
                    Workbook.Main.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: TEST";
                    //giallo
                    Workbook.Main.Shapes.Item("lbTest").Line.Weight = 2f;
                    Workbook.Main.Shapes.Item("lbTest").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 204, 0));
                    Workbook.Main.Shapes.Item("lbTest").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 102));
                    break;
                case Simboli.PROD:
                    Workbook.Main.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: PRODUZIONE";
                    //bianco normale
                    Workbook.Main.Shapes.Item("lbTest").Line.Weight = 0.75f;
                    Workbook.Main.Shapes.Item("lbTest").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                    Workbook.Main.Shapes.Item("lbTest").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 0, 0));
                    Workbook.Main.Shapes.Item("lbTest").Line.ForeColor.Brightness = +0.75f;
                    break;
            }
            Workbook.Main.Shapes.Item("lbTest").Locked = true;

            if (isProt)
                Sheet.Protected = true;
        }
        /// <summary>
        /// Handler per cambiare i label in base alla modifica dello stato del DB.
        /// </summary>
        /// <param name="db">Database interessato</param>
        /// <param name="online">True se il database è online, false altrimenti.</param>
        public static void ChangeStatoDB(PSO.Core.DataBase.NomiDB db, bool online)
        {
            string labelName = "";
            string labelText = "";
            switch (db)
            {
                case PSO.Core.DataBase.NomiDB.SQLSERVER:
                    labelName = "lbSQLServer";
                    labelText = "Database SQL Server: ";
                    break;
                case PSO.Core.DataBase.NomiDB.IMP:
                    labelName = "lbImpianti";
                    labelText = "Database Impianti: ";
                    break;
                case PSO.Core.DataBase.NomiDB.ELSAG:
                    labelName = "lbElsag";
                    labelText = "Database Elsag: ";
                    break;
            }

            var locked = Workbook.Main.ProtectContents;
            if (locked)
                Workbook.Main.Unprotect(Workbook.Password);
            Workbook.Main.Shapes.Item(labelName).TextFrame.Characters().Text = labelText + (online ? "OPERATIVO" : "FUORI SERVIZIO");
            if (online)
            {
                //bianco normale
                Workbook.Main.Shapes.Item(labelName).Line.Weight = 0.75f;
                Workbook.Main.Shapes.Item(labelName).Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                Workbook.Main.Shapes.Item(labelName).Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 0, 0));
                Workbook.Main.Shapes.Item(labelName).Line.ForeColor.Brightness = +0.75f;
            }
            else
            {
                //rosso
                Workbook.Main.Shapes.Item(labelName).Line.Weight = 2f;
                Workbook.Main.Shapes.Item(labelName).Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(140, 56, 54));
                Workbook.Main.Shapes.Item(labelName).Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 80, 77));
            }

            if (locked)
                Workbook.Main.Protect(Workbook.Password);
        }
        /// <summary>
        /// Handler per la modifica del label che indica il mercato attivo.
        /// </summary>
        /// <param name="mercato">La stringa con il nome del mercato.</param>
        public static void ChangeMercatoAttivo(string mercato)
        {
            Workbook.Main.Shapes.Item("lbMercato").Locked = false;
            Workbook.Main.Shapes.Item("lbMercato").TextFrame.Characters().Text = mercato;
            Workbook.Main.Shapes.Item("lbMercato").Locked = true;
        }

        public static void ScriviStagione(int idStagione)
        {
            var sheetPrevisione = Workbook.Sheets.Cast<Excel.Worksheet>()
                    .Where(s => s.Name == "Previsione")
                    .FirstOrDefault();

            bool wasProtected = sheetPrevisione.ProtectContents;
            if (wasProtected)
                sheetPrevisione.Unprotect(Workbook.Password);

            if (sheetPrevisione != null)
            {
                DefinedNames definedNames = new DefinedNames("Previsione");
                DateTime dataFine = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni);
                Range rng = definedNames.Get("CT_TORINO", "STAGIONE", Date.SuffissoDATA1, Date.GetSuffissoOra(1)).Extend(colOffset: Date.GetOreIntervallo(dataFine));
                sheetPrevisione.Range[rng.ToString()].Value = idStagione;
                if (!Simboli.ModificaDati && DataBase.OpenConnection())
                {
                    Handler.StoreEdit(sheetPrevisione, sheetPrevisione.Range[rng.ToString()]);
                    DataBase.SalvaModificheDB();

                    DataBase.CloseConnection();
                }
                
            }
            if (wasProtected)
                sheetPrevisione.Protect(Workbook.Password);
        }

        #endregion
    }
}
