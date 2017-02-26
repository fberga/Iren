
namespace Iren.PSO.Applicazioni
{
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

        protected override void DatiRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.UpdateData();
        }
    }

}
