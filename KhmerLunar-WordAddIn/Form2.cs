using System;
using System.Windows.Forms;

namespace KhmerLunar_WordAddIn
{
    public partial class Form2 : Form
    {
        private MonthCalendar monthCalendar;
        private DateTime selectedDate;

        public Form2()
        {
            InitializeComponent();

            // Initialize calendar
            monthCalendar = new MonthCalendar();
            monthCalendar.MaxSelectionCount = 1;
            monthCalendar.DateSelected += MonthCalendar_DateSelected;

            // Position the calendar
            monthCalendar.Location = new System.Drawing.Point(20, 20);

            // Add calendar to the form
            this.Controls.Add(monthCalendar);
        }

        private void MonthCalendar_DateSelected(object sender, DateRangeEventArgs e)
        {
            selectedDate = e.Start;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (selectedDate != default(DateTime))
            {
                Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Microsoft.Office.Interop.Word.Selection selection = Globals.ThisAddIn.Application.Selection;
                selection.TypeText(KhmerLunar.GetKhmerFullDate(selectedDate));
                this.Close();
            }
            else
            {
                MessageBox.Show("Please select a date.");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
