using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Iren.PSO.Base
{
    public static class Simboli
    {
        public static string LocalBasePath 
        { get { return @"%APPDATA%\PSO\"; } }
        public static string RemoteBasePath
        {   /* Modifica per rliascio in Test ***** BEGIN ***** */
            get { return @"\\srvpso\Applicazioni\PSO_TEST"; } //TODO Riportare a PSO per rilasci in prod
          //   get { return @"\\srvpso\Applicazioni\PSO"; }
            /* Modifica per rliascio in Test ***** END ***** */
        } 

        private readonly static Dictionary<int, string> _fileApplicazione = new Dictionary<int, string>()
        {
            {1, "OfferteMGP"},
            {2, "InvioProgrammi"},
            {3, "InvioProgrammi"},
            {4, "InvioProgrammi"},
            {5, "ProgrammazioneImpianti"},
            {6, "UnitCommitment"},
            {7, "PrezziMSD"},
            {8, "SistemaComandi"},
            {9, "OfferteMSD"},
            {10, "OfferteMB"},
            {11, "ValidazioneTL"},
            {12, "PrevisioneCT"},
            {13, "InvioProgrammi"},
            {14, "ValidazioneGAS"},
            {15, "PrevisioneGAS"},
            //TODO      
            // Modifica per InvioProgrammi MSD5 e MSD6 ***** BEGIN *****
            {16, "InvioProgrammi"},
            {17, "InvioProgrammi"},
            // Modifica per InvioProgrammi MSD5 e MSD6 ***** END *****
            {18, "OfferteMI"}
        };

        public static Dictionary<int, string> FileApplicazione 
        { get { return _fileApplicazione; } }
        
        public const string DEV = "Dev";
        public const string TEST = "Test";
        public const string PROD = "Prod";

        public const string UNION = ".";

        public static string NomeApplicazione 
        { get; set; }
        
        private static bool _emergenzaForzata = false;
        public static bool EmergenzaForzata 
        {
            get
            {
                return _emergenzaForzata;
            }
            set
            {
                if (_emergenzaForzata != value)
                {
                    _emergenzaForzata = value;

                    bool autoCalc = Workbook.Application.Calculation == Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic;

                    if (autoCalc)
                        Workbook.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;

                    bool screenUpdating = Workbook.ScreenUpdating;
                    if (screenUpdating)
                        Workbook.ScreenUpdating = false;

                    bool isProtected = Workbook.Main.ProtectContents;
                    if (isProtected)
                        Workbook.Main.Unprotect(Workbook.Password);

                    Riepilogo main = new Riepilogo(Workbook.Main);
                    if (value)
                        main.RiepilogoInEmergenza();
                    else
                        if (DataBase.OpenConnection())
                        {
                            main.UpdateData();
                            DataBase.CloseConnection();
                        }

                    Workbook.AggiornaLabelStatoDB();

                    if (isProtected)
                        Workbook.Main.Protect(Workbook.Password);

                    if(autoCalc)
                        Workbook.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic;

                    if (screenUpdating)
                        Workbook.ScreenUpdating = true;
                }
            }
        }

        private static bool _modificaDati = false;
        public static bool ModificaDati 
        { 
            get 
            { 
                return _modificaDati; 
            } 
            
            set 
            {
                _modificaDati = value;
                Handler.ChangeModificaDati(_modificaDati);
            }
        }

        private static bool _sqlServerOnline = true;
        public static bool SQLServerOnline
        {
            get
            {
                return _sqlServerOnline;
            }

            set
            {
                _sqlServerOnline = value;
                Handler.ChangeStatoDB(Core.DataBase.NomiDB.SQLSERVER, _sqlServerOnline);
            }
        }

        private static bool _impiantiOnline = true;
        public static bool ImpiantiOnline
        {
            get
            {
                return _impiantiOnline;
            }

            set
            {
                _impiantiOnline = value;
                Handler.ChangeStatoDB(Core.DataBase.NomiDB.IMP, _impiantiOnline);
            }
        }

        private static bool _elsagOnline = true;
        public static bool ElsagOnline
        {
            get
            {
                return _elsagOnline;
            }

            set
            {
                _elsagOnline = value;
                Handler.ChangeStatoDB(Core.DataBase.NomiDB.ELSAG, _elsagOnline);
            }
        }
        
        public static int[] rgbSfondo = { 228, 144, 144 };
        public static int[] rgbLinee = { 176, 0, 0 };
        public static int[] rgbTitolo = { 206, 58, 58 };

        private readonly static Dictionary<int, string> oreMSD = new Dictionary<int, string>() 
        { 
            //TODO modificare in base ai nuovi orari
             // Modifica per InvioProgrammi MSD5 e MSD6 e nuovi orari ***** BEGIN *****
            {0, "MSD2"},
            {1, "MSD2"},
            {2, "MSD2"},
            {3, "MSD2"},
            {4, "MSD3"},
            {5, "MSD3"},
            {6, "MSD3"},
            {7, "MSD3"},
            {8, "MSD4"},
            {9, "MSD4"},
            {10, "MSD4"},
            {11, "MSD4"},
            {12, "MSD5"},
            {13, "MSD5"},
            {14, "MSD5"},
            {15, "MSD5"},
            {16, "MSD6"},
            {17, "MSD6"},
            {18, "MSD6"},
            {19, "MSD6"},
            {20, "MSD1"},
            {21, "MSD1"},
            {22, "MSD1"},
            {23, "MSD1"},
            // Modifica per InvioProgrammi MSD5 e MSD6 e nuovi orari ***** END *****

             /*
             // Orari fino al 01/02/2017
            {0, "MSD1"},
            {1, "MSD1"},
            {2, "MSD1"},
            {3, "MSD1"},
            {4, "MSD2"},
            {5, "MSD2"},
            {6, "MSD2"},
            {7, "MSD2"},
            {8, "MSD3"},
            {9, "MSD3"},
            {10, "MSD3"},
            {11, "MSD3"},
            {12, "MSD4"},
            {13, "MSD4"},
            {14, "MSD4"},
            {15, "MSD4"},
            {16, "MSD4"},
            {17, "MSD4"},
            {18, "MSD4"},
            {19, "MSD1"},
            {20, "MSD1"},
            {21, "MSD1"},
            {22, "MSD1"},
            {23, "MSD1"},
              */
        };

        public static Dictionary<int, string> OreMSD
        { get { return oreMSD; } }

        public static string GetMercatoPrec(string mercato)
        {
            var index = Workbook.Repository[DataBase.TAB.MERCATI].AsEnumerable()
                .Where(r => r["DesMercato"].Equals(mercato))
                .Select(r => Workbook.Repository[DataBase.TAB.MERCATI].Rows.IndexOf(r))
                .FirstOrDefault();

            if (index > 0)
                return Workbook.Repository[DataBase.TAB.MERCATI].Rows[index - 1]["DesMercato"].ToString();

            return "";

        }

        private readonly static Dictionary<string, SpecMercato> mercatiMB = new Dictionary<string, SpecMercato>()
        {
            //TODO modificare in base a nuovi orari
            /*
            {"MB1", new MB(0,1,8)},
            {"MB2", new MB(7,9,12)},
            {"MB3", new MB(11,13,16)},
            {"MB4", new MB(15,17,22)},
            {"MB5", new MB(21,23,25)}
            */
            /******************** Modifica nuovi mercati MB  BEGIN ********************/
            {"MB1", new SpecMercato(0,1,4)}, // Da modificare
            {"MB2", new SpecMercato(3,5,25)},
            {"MB3", new SpecMercato(7,9,25)},
            {"MB4", new SpecMercato(11,13,25)},
            {"MB5", new SpecMercato(15,17,25)},
            {"MB6", new SpecMercato(19,21,25)}
            /******************** Modifica nuovi mercati MB  END ********************/
        };

        //06/02/2017 MOD: aggiunta orari mercati MI
        /*
        private readonly static Dictionary<string, SpecMercato> mercatiMI = new Dictionary<string, SpecMercato>()
        {
            //TODO controllo orari effettivi
            
            //{"MI1", new SpecMercato(15,1,24)}, //boohh
            //{"MI2", new SpecMercato(16,1,4)},
            //{"MI3", new SpecMercato(24,5,8)},
            //{"MI4", new SpecMercato(3,9,12)},
            //{"MI5", new SpecMercato(8,13,16)},
            //{"MI6", new SpecMercato(11,17,20)},
            //{"MI7", new SpecMercato(15,21,25)}
            
            {"MI1", new SpecMercato(15,1,25)}, //boohh
            {"MI2", new SpecMercato(16,1,25)},
            {"MI3", new SpecMercato(20,5,25)},
            {"MI4", new SpecMercato(4,9,25)},
            {"MI5", new SpecMercato(8,13,25)},
            {"MI6", new SpecMercato(11,17,25)},
            {"MI7", new SpecMercato(13,21,25)}
        };
        */

        private readonly static Dictionary<string, bool> mercatiMI_infoDay = new Dictionary<string, bool>()
        {
            {"MI1", true}, //boohh
            {"MI2", true},
            {"MI3", true},
            {"MI4", true},
            {"MI5", false},
            {"MI6", false},
            {"MI7", false}
        };


        /* INIZIO */
        /* Dizionario utilizzando il nuovo costruttore */
        /* Tupla contenente  | mercato | Inizio offerte | Fine offerte | Inizio oggetto mercato | Riferimento a giorno successivo | */
        private readonly static List<Tuple<string, TimeSpan, TimeSpan, int, bool>> mercatiMI = new List<Tuple<string, TimeSpan, TimeSpan, int, bool>>
        {
            Tuple.Create("MI1", new TimeSpan(12,55,0), new TimeSpan(15,0,0), 1, true ),
            Tuple.Create("MI2", new TimeSpan(12,55,0), new TimeSpan(16,30,0), 1, true ),
            Tuple.Create("MI3", new TimeSpan(16,30,0), new TimeSpan(23,45,0), 5, true ),
            Tuple.Create("MI4", new TimeSpan(0,0,0), new TimeSpan(3,45,0), 9, false),
            Tuple.Create("MI5", new TimeSpan(0,0,0), new TimeSpan(7,45,0), 13, false),
            Tuple.Create("MI6", new TimeSpan(0,0,0), new TimeSpan(11,15,0), 17, false),
            Tuple.Create("MI7", new TimeSpan(0,0,0), new TimeSpan(15,45,0), 21, false),
        };

        /* Tupla di riferimento per la selezione del mercato all'avvio. Contiene: | mercato | Inizio fascia oraria | Fine fascia oraria */
        /* PS: si gestisce la fascia oraria non contigua per il mercato MI4  */
        private readonly static List<Tuple<string, TimeSpan, TimeSpan>> marketTimeDefinition = new List<Tuple<string, TimeSpan, TimeSpan>>
        {
            Tuple.Create("MI1", new TimeSpan(13,0,0), new TimeSpan(15,0,0)),
            Tuple.Create("MI2", new TimeSpan(15,0,0), new TimeSpan(16,30,0)),
            Tuple.Create("MI3", new TimeSpan(16,30,0), new TimeSpan(20,0,0)),
            Tuple.Create("MI4", new TimeSpan(20,0,0), new TimeSpan(24,0,0)),
            Tuple.Create("MI4", new TimeSpan(0,0,0), new TimeSpan(4,0,0)),
            Tuple.Create("MI5", new TimeSpan(4,0,0), new TimeSpan(7,45,0)),
            Tuple.Create("MI6", new TimeSpan(7,45,0), new TimeSpan(11,15,0)),
            Tuple.Create("MI7", new TimeSpan(11,15,0), new TimeSpan(13,0,0))
        };

        /* FINE */



        public static Dictionary<string, SpecMercato> MercatiMB { get { return mercatiMB; } }
        public static List<Tuple<string, TimeSpan, TimeSpan, int, bool>> MercatiMI { get { return mercatiMI; } }

        // nuovo dictionary
        //public static Dictionary<string, NewSpecMercato> MercatiMI_v2 { get { return mercatiMI_v2; } }

        //08/02/2017 MOD: aggiunto metodo che restituisce il nome del mercato. TODO mettere tutta la gestione mercati sul DB
        public static string GetActiveMarket(int hour)
        {
            string result = "";
            /*09/02/2017 inserito nuovo metodo GetActiveMarkets che estrae la lista dei mercati attivi per gestione popolazione combo lista mercati*/
            if (Workbook.IdApplicazione == 10)
            {
                result = GetActiveMarkets(hour).FirstOrDefault();
            }
            else if (Workbook.IdApplicazione == 18)
            {
                TimeSpan freeze = DateTime.Now.TimeOfDay;
                Tuple<string, TimeSpan, TimeSpan> market = marketTimeDefinition.Where(x => x.Item2 <= freeze && x.Item3 > freeze).FirstOrDefault();

                result = market.Item1;
            }

            return result;
            /*
            if (Workbook.IdApplicazione == 10)
            {
                mercato = Simboli.MercatiMB
                    .Where(kv => kv.Value.Chiusura > hour)
                    .Select(kv => kv.Key)
                    .FirstOrDefault();
            }
            else if (Workbook.IdApplicazione == 18)
            {
                   
                if (hour >= Simboli.MercatiMI["MI7"].Chiusura)
                {
                    mercato = Simboli.MercatiMI
                        .Where(kv => kv.Value.Chiusura > hour)
                        .Select(kv => kv.Key)
                        .FirstOrDefault();
                }
                else 
                {
                    mercato = Simboli.MercatiMI
                        .Where(kv => (kv.Value.Chiusura > hour && !(kv.Key == "MI1" || kv.Key == "MI2" || kv.Key == "MI3")))
                        .Select(kv => kv.Key)
                        .FirstOrDefault();
                }
            }
            return mercato;
             */
        }

        public static List<string> GetActiveMarkets(int hour)
        {
            /*09/02/2017 inserito nuovo metodo GetActiveMarkets che estrae la lista dei mercati attivi per gestione popolazione combo lista mercati*/
            List<string> result = new List<string>();
            
            if (Workbook.IdApplicazione == 10)
            {
                result = Simboli.MercatiMB
                    .Where(kv => kv.Value.Chiusura > hour)
                    .Select(kv => kv.Key).ToList();
            }
            else if (Workbook.IdApplicazione == 18)
            {
                //Simboli.MercatiMI .Select(x => x.Key).ToList();
                result = mercatiMI.Select(x => x.Item1).ToList(); 
                /*
                if (hour >= Simboli.MercatiMI["MI7"].Chiusura)
                {
                    result = Simboli.MercatiMI
                        .Where(kv => kv.Value.Chiusura > hour)
                        .Select(kv => kv.Key).ToList();
                }
                else
                {
                    result = Simboli.MercatiMI
                        .Where(kv => (kv.Value.Chiusura > hour && !(kv.Key == "MI1" || kv.Key == "MI2" || kv.Key == "MI3")))
                        .Select(kv => kv.Key).ToList();
                }
                */
            }
            return result;
        } 

        //08/02/2017 MOD: spostata logica di ricerca mercato
        //public static int GetMarketOffset(int hour)
        //{
        //    if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1"))
        //    {
        //        //01/02/2017 FIX: Nuova logica mercati
        //        //int offset = Simboli.MercatiMB["MB1"].Fine;
        //        //if (hour >= Simboli.MercatiMB["MB2"].Chiusura)
        //        //{
        //        //    string mercatoChiuso = Simboli.MercatiMB
        //        //        .Where(kv => kv.Value.Chiusura <= hour)
        //        //        .Select(kv => kv.Key)
        //        //        .Last();
        //        //    offset = Simboli.MercatiMB[mercatoChiuso].Fine;
        //        //}
        //        //06/02/2017 MOD: distinzione tra MB e MI
        //        int offset = 0;
        //        if (Workbook.IdApplicazione == 10)
        //        {
        //            offset = Simboli.MercatiMB["MB1"].Fine;
        //            if (hour >= Simboli.MercatiMB["MB2"].Chiusura)
        //            {
        //                string primoMercatoAperto = Simboli.MercatiMB
        //                    .Where(kv => kv.Value.Chiusura > hour)
        //                    .Select(kv => kv.Key)
        //                    .FirstOrDefault();
        //                if (primoMercatoAperto == null)
        //                    offset = Date.GetOreGiorno(Workbook.DataAttiva);
        //                else
        //                    offset = Simboli.MercatiMB[primoMercatoAperto].Inizio - 1;
        //            }
        //        }
        //        else if (Workbook.IdApplicazione == 18)
        //        {
        //            /* //TODO rivedere logica
        //             * esempio: ore 23:20
        //             *  - sono in MI3 e lavoro su D + 1
        //             *  - l'applicativo è posizionato correttamente su D + 1
        //             *  - le ore da visualizzare sbloccate sono (credo) da 5 in avanti
        //             *  - 
        //             */
    //     //       offset = Simboli.MercatiMI["MI1"].Fine;
    //            string primoMercatoAperto = "";
                   

    //            if (hour >= Simboli.MercatiMI["MI7"].Chiusura)
    //            {
    //                primoMercatoAperto = Simboli.MercatiMI
    //                   .Where(kv => kv.Value.Chiusura > hour)
    //                   .Select(kv => kv.Key)
    //                   .FirstOrDefault();
    //                if (primoMercatoAperto == null)
    //                    offset = Date.GetOreGiorno(Workbook.DataAttiva);
    //                else
    //                    offset = Simboli.MercatiMI[primoMercatoAperto].Inizio - 1;
    //            }
    //            else 
    //            {
    //                 primoMercatoAperto = Simboli.MercatiMI
    //                   .Where(kv => (kv.Value.Chiusura > hour && !(kv.Key == "MI1" || kv.Key == "MI2" || kv.Key == "MI3")))
    //                   .Select(kv => kv.Key)
    //                   .FirstOrDefault();
    //                if (primoMercatoAperto == null)
    //                    offset = Date.GetOreGiorno(Workbook.DataAttiva);
    //                else
    //                    offset = Simboli.MercatiMI[primoMercatoAperto].Inizio - 1;
    //            }
    //        }

    //        return offset;
    //    }
    //    return 0;
    //}
        
        public static int GetMarketOffset(int hour)
        {
            if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1") )
            {
                //01/02/2017 FIX: Nuova logica mercati
                //int offset = Simboli.MercatiMB["MB1"].Fine;
                //if (hour >= Simboli.MercatiMB["MB2"].Chiusura)
                //{
                //    string mercatoChiuso = Simboli.MercatiMB
                //        .Where(kv => kv.Value.Chiusura <= hour)
                //        .Select(kv => kv.Key)
                //        .Last();
                //    offset = Simboli.MercatiMB[mercatoChiuso].Fine;
                //}
                //06/02/2017 MOD: distinzione tra MB e MI
                int offset = 0;
                if (Workbook.IdApplicazione == 10)
                {
                    offset = Simboli.MercatiMB["MB1"].Fine;
                    if (hour >= Simboli.MercatiMB["MB2"].Chiusura)
                    {
                        string primoMercatoAperto = GetActiveMarket(hour);
                        if (primoMercatoAperto == null)
                            offset = Date.GetOreGiorno(Workbook.DataAttiva);
                        else
                            offset = Simboli.MercatiMB[primoMercatoAperto].Inizio - 1;
                    }
                }
                // Non verra mai chiamata con id = 18
                /*
                else if (Workbook.IdApplicazione == 18)
                {
                     //TODO rivedere logica
                     //esempio: ore 23:20
                     //  - sono in MI3 e lavoro su D + 1
                     //  - l'applicativo è posizionato correttamente su D + 1
                     //  - le ore da visualizzare sbloccate sono (credo) da 5 in avanti
                     //  - 
                     
                    string primoMercatoAperto = GetActiveMarket(hour);

                    if (primoMercatoAperto == null)
                        offset = Date.GetOreGiorno(Workbook.DataAttiva);
                    else
                        offset = Simboli.MercatiMI[primoMercatoAperto].Inizio - 1;
                }
                */
                return offset;
            }
            return 0;
        }

        // utile per l'abilitazione/disabilitazione delle celle dei fogli
        public static int GetMarketOffsetMI(string mercato, DateTime dataExcel)
        {
            DateTime freeze = DateTime.Now;
            if (dataExcel.Date == freeze.Date)
            {
                var tmp = mercatiMI.Where(x => !x.Item5 && x.Item1.Equals(mercato) && x.Item2 <= freeze.TimeOfDay && freeze.TimeOfDay < x.Item3).FirstOrDefault();
                return tmp != null ? tmp.Item4 : 25;
            }
            else if (dataExcel.Date == DateTime.Now.Date.AddDays(1))
            {
                var tmp = mercatiMI.Where(x => /*x.Item5 &&*/ x.Item1.Equals(mercato)).FirstOrDefault();
                return tmp != null ? tmp.Item4 : 25;
            }

            return 25;

            /*
            if (dataExcel.Date > DateTime.Now.Date)
            {
                if (mercatiMI_infoDay[mercato])
                    return MercatiMI[mercato].Inizio;
                else
                    return 25;
            }
            else if(dataExcel.Date == DateTime.Now.Date)
            {
                if (mercatiMI_infoDay[mercato])
                    return 25;
                else
                    return MercatiMI[mercato].Inizio;
            }

            
            */
        }

        public static Range GetMarketCompleteRange(string mercato, DateTime giorno, Range rng)
        {
            if (!mercatiMB.ContainsKey(mercato))
                return null;

            int[] orario = new int[2] { Simboli.MercatiMB[mercato].Inizio, Math.Min(Simboli.MercatiMB[mercato].Fine, Date.GetOreGiorno(giorno)) };

            return new Range(rng.StartRow, rng.StartColumn + orario[0] - 1, 1, orario[1] - orario[0] + 1);
        }
    }
}
//06/02/2017 MOD: Cambiato nome per utilizzare con MI
//public class MB
public class SpecMercato
{
    public int Chiusura { get; private set; }
    public int Inizio { get; private set; }
    public int Fine { get; private set; }

    public string ToString()
    {
        return "C: " + Chiusura + "; I: " + Inizio + "; F: " + Fine;
    }
    public SpecMercato(int chiusura, int inizio, int fine)
    {
        Chiusura = chiusura;
        Inizio = inizio;
        Fine = fine;
    }


    
    

    
}
/* INIZIO */
/* G.U. Nuova Classe */
public class NewSpecMercato
{
    /* INIZIO */ //mi piaceee con i timespan!!
    /* G.U. nuovo codice per sostituzione definizione SpecMercato aggiungendo l'inizio dell'offerta e utilizzando i TimeSpan
       per affinare la definizione degli intervalli di attività */
    public TimeSpan InizioOfferta { get; private set; }
    public TimeSpan FineOfferta { get; private set; }
    public int InizioMercato { get; private set; }
    public int FineMercato { get; private set; }
    public bool DayAfter { get; private set; }
    /* FINE */

    public string ToString()
    {
        return "FO: " + FineOfferta.Hours + ":" + FineOfferta.Minutes + "; IO: " + InizioOfferta.Hours + ":" + InizioOfferta.Minutes + "; IM: " + InizioMercato + "; FM: " + FineMercato + "; DA: " + DayAfter.ToString();
    }

    public NewSpecMercato(TimeSpan inizioOfferta, TimeSpan fineOfferta, int inizioMercato, int fineMercato, bool dayAfter)
    {
        InizioOfferta = inizioOfferta;
        FineOfferta = fineOfferta;
        InizioMercato = inizioMercato;
        FineMercato = fineMercato;
        DayAfter = dayAfter;
    }
    
}

/* FINE */
