using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TalentWindowsFormsApp
{
    public partial class CalendarForm : Form
    {
        public string CalenderValue { get; set; }
        public CalendarForm()
        {
            InitializeComponent();
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            CalenderValue = monthCalendar1.SelectionStart.ToString("yyyy/MM/dd");
            this.DialogResult = DialogResult.OK;
            this.Hide();
        }
    }
}
