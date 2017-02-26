using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    class DisabilitaMenuStrip : ContextMenuStrip
    {
        private ToolStripItem _disabilita;
        private ToolStripItem _abilita;

        public DisabilitaMenuStrip()
        {
            _disabilita = Items.Add("Disabilita");
            _abilita = Items.Add("Abilita");
        }

        protected override void OnOpening(System.ComponentModel.CancelEventArgs e)
        {
            IRibbonControl ctrl = SourceControl as IRibbonControl;

            if(ctrl != null)
            {
                _abilita.Enabled = !ctrl.Enabled;
                _disabilita.Enabled = ctrl.Enabled;
            }
            else
            {
                e.Cancel = true;
            }

            base.OnOpening(e);
        }

        protected override void OnItemClicked(ToolStripItemClickedEventArgs e)
        {
            IRibbonControl ctrl = SourceControl as IRibbonControl;

            if (e.ClickedItem == _disabilita)
            {
                ctrl.Enabled = false;
                _disabilita.Enabled = false;
                _abilita.Enabled = true;
            }
            else
            {
                ctrl.Enabled = true;
                _disabilita.Enabled = true;
                _abilita.Enabled = false;
            }
            
            
            base.OnItemClicked(e);
        }

    }
}
