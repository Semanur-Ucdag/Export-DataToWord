using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace ExportDataToWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            var word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true;
            word.WindowState = WdWindowState.wdWindowStateNormal;
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Add();
            Microsoft.Office.Interop.Word.Paragraph paragraph = doc.Paragraphs.Add();
            paragraph.Range.Text = richTextBox1.Text;
            string folderPath = @"D:\06072023SemanurBackupFolder\source\repos\WordeAktarma\ExportDataToWord";
            string filePath = System.IO.Path.Combine(folderPath, "test.docx");
            doc.SaveAs2(filePath);
            Process.Start(filePath);



            //doc.SaveAs2("WordFile.docx");
            //Process.Start("WordFile.docx");







        }
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

       
    }
}
