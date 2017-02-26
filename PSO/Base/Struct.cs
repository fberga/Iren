
namespace Iren.PSO.Base
{
    public class Struct
    {
        #region Strutture

        public struct Cella
        {
            public struct Width
            {
                public double empty,
                    dato,
                    entita,
                    informazione,
                    unitaMisura,
                    parametro,
                    jolly1,
                    riepilogo;
            }
            public struct Height
            {
                public double normal,
                    empty;
            }

            public Width width;
            public Height height;
        }

        #endregion

        #region Costanti

        public const int COLORE_VARIAZIONE_POSITIVA = 4;
        public const int COLORE_VARIAZIONE_NEGATIVA = 38;

        #endregion

        #region Variabili

        public static string tipoVisualizzazione = "O";
        public static int intervalloGiorni = 0;
        public static bool visualizzaRiepilogo = true;
        public static Cella cell;
        public static bool visLinkEntita = true;

        public int numRigheMenu = 1;
        public int numEleMenu = 1;

        public int colBlock = 5,
            rigaBlock = 6,
            rigaGoto = 3,
            colRecap = 165,
            rowRecap = 2;
        public bool visData0H24 = false,
            //visParametro = false,
            visSelezione = false;

        #endregion
    }
}
