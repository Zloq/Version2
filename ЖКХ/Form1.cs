using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;

namespace ЖКХ
{
    public partial class Form1 : Form
    {
        OpenFileDialog ofd = new OpenFileDialog();
        int num = 0;
        int min;
        int sec;
        public Form1()
        {
            InitializeComponent();
            textBox6.Text = "Стоимость 1 кубического метра холодной воды стоит 9,32 рублей," +
                "\n 1 кубический метр горячей воды – 12,76 \nрублей";
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// Рассчет данных по тарифам
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox1.Text == "" || textBox2.Text == "" || checkBox1.Checked == false || textBox3.Text == "")
                {
                    MessageBox.Show("Заполните все данные", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    double Hot = Convert.ToDouble(textBox1.Text);
                    double Cold = Convert.ToDouble(textBox2.Text);
                    double Light = Convert.ToDouble(textBox3.Text);
                    double ExistingIndications = Convert.ToDouble(textBox5.Text);
                    double PastTestimony = Convert.ToDouble(textBox4.Text);

                    double WaterH = Convert.ToDouble(Hot * Price.HotWater);
                    double WaterC = Convert.ToDouble(Cold * Price.ColdWater);
                    double LightV = Convert.ToDouble(ExistingIndications - PastTestimony);
                    double LightT = Convert.ToDouble(LightV * Light);
                    double PricetotalHCHL = WaterH + WaterC + Price.Heating + LightT;
                    textBox6.Text = Convert.ToString($"{PricetotalHCHL}.Руб");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        /// <summary>
        /// Запрет на прописть букв, запрещенно все кроме цифр
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(Char.IsNumber(e.KeyChar) | e.KeyChar == '\b') return;
            else
                e.Handled = true;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | e.KeyChar == '\b') return;
            else
                e.Handled = true;
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
           
        }

        /// <summary>
        /// Добавление картинки и вывод ее на форму 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            ofd.Filter = "Image Files(*.JPG;*.JPEG;)|*.JPG;*.JPEG; | All files(*.*) | *.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    pictureBox1.Image = new Bitmap(ofd.FileName);
                }
                catch
                {
                    MessageBox.Show("Невозможно открыть выбранный файл", "Ошибка");
                }
            }
        }
        /// <summary>
        /// Создание и формирование чека
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            //Создаём объект документа
            Word.Document doc = null;
            try
            {
                // Создаём объект приложения
                Word.Application app = new Word.Application();
                // Путь до шаблона документа
                string source = Path.Combine(Directory.GetCurrentDirectory(), "Чек.docx");
                // Открываем
                doc = app.Documents.Add(source);
                doc.Activate();

                // Добавляем информацию
                // wBookmarks содержит все закладки
                Word.Bookmarks wBookmarks = doc.Bookmarks;
                Word.Range wRange;
                int i = 0;
                num++;
                string nm = Convert.ToString(num);
                string[] data = new string[3] {DateTime.Now.ToShortDateString(), nm, textBox6.Text};
                foreach (Word.Bookmark mark in wBookmarks)
                {

                    wRange = mark.Range;
                    wRange.Text = data[i];
                    i++;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
    }
}