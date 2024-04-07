using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Veterinar
{
    public partial class LookAppointmentForm : Form
    {
        string date, client, pet, vaccine, service, whathurt, whatwasdone, whatneedtodo;

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            // Создать шрифт myFont
            Font myFont = new Font("Arial", 14, FontStyle.Regular, GraphicsUnit.Pixel);

            string curLine; // текущая выводимая строка

            // Отступы внутри страницы
            float leftMargin = e.MarginBounds.Left; // отступы слева в документе
            float topMargin = e.MarginBounds.Top; // отступы сверху в документе
            float yPos = 0; // текущая позиция Y для вывода строки

            int nPages; // количество страниц
            int nLines; // максимально-возможное количество строк на странице
            int i; // номер текущей строки для вывода на странице

            // Вычислить максимально возможное количество строк на странице
            nLines = (int)(e.MarginBounds.Height / myFont.GetHeight(e.Graphics));

            // Вычислить количество страниц для печати
            nPages = (richTextBox1.Lines.Length - 1) / nLines + 1;

            // Цикл печати/вывода одной страницы
            i = 0;
            while ((i < nLines) && (counter < richTextBox1.Lines.Length))
            {
                // Взять строку для вывода из richTextBox1
                curLine = richTextBox1.Lines[counter];

                // Вычислить текущую позицию по оси Y
                yPos = topMargin + i * myFont.GetHeight(e.Graphics);
                // Вывести строку в документ
                e.Graphics.DrawString(curLine, myFont, Brushes.Black,
                  leftMargin, yPos, new StringFormat());

                counter++;
                i++;
            }

            // Если весь текст не помещается на 1 страницу, то
            // нужно добавить дополнительную страницу для печати
            e.HasMorePages = false;

            if (curPage < nPages)
            {
                curPage++;
                e.HasMorePages = true;
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            pageSetupDialog1.ShowDialog();
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.DefaultExt = "*.txt";
            sfd.Filter = "Text files|*.txt";
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK &&
                sfd.FileName.Length > 0)
            {
                using (StreamWriter sw = new StreamWriter(sfd.FileName, true))
                {
                    sw.WriteLine(label1.Text);
                    sw.WriteLine(label2.Text);
                    sw.WriteLine(label3.Text);
                    sw.WriteLine(label4.Text);
                    sw.WriteLine(richTextBox1.Text);
                    sw.Close();
                }
            }
        }

        private void printDocument1_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            counter = 0;
            curPage = 1;
        }

        int counter = 0; // сквозной номер строки в массиве строк, которые выводятся
        int curPage; // текущая страница

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            printPreviewDialog1.ShowDialog();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (printDialog1.ShowDialog() == DialogResult.OK)
                printDocument1.Print();
        }

        private void LookAppointmentForm_Load(object sender, EventArgs e)
        {
            this.label1.Text = date;
            this.label2.Text = client;
            this.label3.Text = pet;
            this.label4.Text = vaccine;
            this.textBox1.Text = service;
            this.textBox2.Text = whathurt;
            this.textBox3.Text = whatwasdone;
            this.textBox4.Text = whatneedtodo;
            this.richTextBox1.Text = "Услуга: " + service + "\n" + "Что болело: " + whathurt + "\n" + "Что было сделано: " + "\n"
                + whatwasdone + "\n" + "Что нужно сделать: " + whatneedtodo;

            string str = this.richTextBox1.Text;

            String[] sublines = str.Split(' ');
            str = null;
            int length = 80;
            int j = 0;
            for (int i = 0; i < sublines.Count(); i++)
            {
                if (j + sublines[i].Length < length)
                {
                    str = str + sublines[i] + " ";
                    j = j + sublines[i].Length;
                }
                else
                {
                    j = 0;
                    str = str + "\r\n";
                    i--;
                }
            }
            this.richTextBox1.Text = this.label1.Text + "\n" + this.label2.Text + "\n" + this.label3.Text + "\n" 
                + this.label4.Text + "\n" + str;
        }
        public LookAppointmentForm(string date, string client, string pet, string vaccine, string service, string whathurt, string whatwasdone, string whatneedtodo)
        {
            this.date = date;
            this.client = client;
            this.pet = pet;
            this.vaccine = vaccine;
            this.service = service;
            this.whathurt = whathurt;
            this.whatwasdone = whatwasdone;
            this.whatneedtodo = whatneedtodo;
            InitializeComponent();
        }
    }
    public static class Extensions
    {
        public static IEnumerable<string> Split(this string str, int n)
        {
            if (String.IsNullOrEmpty(str) || n < 1)
            {
                throw new ArgumentException();
            }
            for (int i = 0; i < str.Length; i += n)
            {
                yield return str.Substring(i, Math.Min(n, str.Length - i));
            }
        }
    }
}
