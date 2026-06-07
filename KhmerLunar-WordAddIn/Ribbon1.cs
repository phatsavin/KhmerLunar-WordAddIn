using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace KhmerLunar_WordAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Interop.Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            selection.TypeText(KhmerLunar.getKhmerLunarString(DateTime.Now));
            //KhmerLunar.getKhmerLunarString(DateTime.Now);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 form = new Form1(); // Create an instance of Form2
            form.ShowDialog();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Interop.Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            selection.TypeText(KhmerLunar.GetKhmerFullDate(DateTime.Now));
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Form2 form = new Form2(); // Create an instance of Form2
            form.ShowDialog();
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Form3 form = new Form3(); // Create an instance of Form2
            form.ShowDialog();
        }
    }
}
