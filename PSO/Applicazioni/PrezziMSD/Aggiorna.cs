
namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Label riepilogo custom.
    /// </summary>
    public class Aggiorna : Base.Aggiorna
    {
        public Aggiorna()
            : base()
        {

        }        

        protected override void StrutturaRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.LoadStructure();
        }
    }

}
