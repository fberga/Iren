using Iren.PSO.Base;
using System;
using System.Windows.Forms;

namespace Iren.PSO.Forms
{
    public partial class FormCalendar : Form
    {
        public FormCalendar()
        {
            InitializeComponent();
            //Application.EnableVisualStyles();
            calObj.SetDate(Workbook.DataAttiva);
            this.Text = Simboli.NomeApplicazione + " - Calendar";
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = System.Windows.Forms.DialogResult.OK;
            Close();
        }

        private void btnANNULLA_Click(object sender, EventArgs e)
        {
            DialogResult = System.Windows.Forms.DialogResult.Cancel;
            Close();
        }

        /// <summary>
        /// Override del metodo ShowDialog di Windows Forms. Restituisce l'oggetto data selezionato se l'utente ha cambiato la selezione, null altrimenti.
        /// </summary>
        /// <returns>Restituisce l'oggetto data selezionato se l'utente ha cambiato la selezione, null altrimenti.</returns>
        public new DateTime ShowDialog()
        {
            base.ShowDialog();
            if (DialogResult == System.Windows.Forms.DialogResult.OK)
                return calObj.SelectionStart;
            else
                return Workbook.DataAttiva;
        }
    }
}
