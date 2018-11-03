using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MyHotelApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Microsoft.Office.Interop.Excel.Application ObjExcel;
        Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

        private void WriteToExcel()
        {
            ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            string NameExcel = @"" + textBox4.Text + "\\" + textBox10.Text + ".xlsx";
            try
            {
                ObjWorkBook = ObjExcel.Workbooks.Open(NameExcel,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
            }
            catch
            {
                MessageBox.Show("Ошибка при открытии файла", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            ObjWorkSheet.Protect(
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            true, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            int i;
            for (i = 1; ; i++)
            {
                if (Convert.ToString(ObjWorkSheet.Cells[i, 1].Value2) == null) break;
                else continue;
            }
            string number = Convert.ToString(i - 1);
            if (String.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Нужно заполнить поле \"Фамилия\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (textBox1.BackColor == Color.LightCoral)
            {
                MessageBox.Show("Неверный формат данных поля \"Фамилия\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (String.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Нужно заполнить поле \"Имя\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (textBox2.BackColor == Color.LightCoral)
            {
                MessageBox.Show("Неверный формат данных поля \"Имя\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (String.IsNullOrEmpty(textBox5.Text))
            {
                MessageBox.Show("Нужно заполнить поле \"Номер комнаты\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (textBox5.BackColor == Color.LightCoral)
            {
                MessageBox.Show("Неверный формат данных поля \"Номер комнаты\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (String.IsNullOrEmpty(textBox3.Text))
            {
                MessageBox.Show("Нужно заполнить поле \"Дата заселения\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (textBox3.BackColor == Color.LightCoral)
            {
                MessageBox.Show("Неверный формат данных поля \"Дата заселения\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (String.IsNullOrEmpty(textBox6.Text))
            {
                MessageBox.Show("Нужно заполнить поле \"Дата внесения проплаты\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (textBox6.BackColor == Color.LightCoral)
            {
                MessageBox.Show("Неверный формат данных поля \"Дата внесения проплаты\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (String.IsNullOrEmpty(textBox7.Text))
            {
                MessageBox.Show("Нужно заполнить поле \"Количество дней\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (textBox7.BackColor == Color.LightCoral)
            {
                MessageBox.Show("Неверный формат данных поля \"Количество дней\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (String.IsNullOrEmpty(textBox8.Text))
            {
                MessageBox.Show("Нужно заполнить поле \"Дата оплаты\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (textBox8.BackColor == Color.LightCoral)
            {
                MessageBox.Show("Неверный формат данных поля \"Дата оплаты\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (String.IsNullOrEmpty(textBox9.Text))
            {
                MessageBox.Show("Нужно заполнить поле \"Сумма\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            if (textBox9.BackColor == Color.LightCoral)
            {
                MessageBox.Show("Неверный формат данных поля \"Сумма\"", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseFileExcel();
                return;
            }
            ObjWorkSheet.Cells[i, 1] = number;
            ObjWorkSheet.Cells[i, 2] = textBox1.Text + " " + textBox2.Text;
            ObjWorkSheet.Cells[i, 3] = textBox5.Text;
            ObjWorkSheet.Cells[i, 4] = comboBox1.Text;
            ObjWorkSheet.Cells[i, 5] = textBox3.Text;
            ObjWorkSheet.Cells[i, 6] = textBox6.Text;
            ObjWorkSheet.Cells[i, 7] = textBox7.Text;
            ObjWorkSheet.Cells[i, 8] = textBox8.Text;
            ObjWorkSheet.Cells[i, 9] = textBox9.Text;
            CloseFileExcel();
            MessageBox.Show("Посетитель успешно добавлен!\n" +
                "Данные о посетителе:\n" +
                "Фамилия: " + textBox1.Text +
                "\nИмя: " + textBox2.Text +
                "\nНомер комнаты: " + textBox5.Text +
                "\nМесто: " + comboBox1.Text +
                "\nДата заселения: " + textBox3.Text +
                "\nДата внесения проплаты: " + textBox6.Text +
                "\nКоличество дней: " + textBox7.Text +
                "\nДата оплаты: " + textBox8.Text +
                "\nСумма: " + textBox9.Text);
            reset(textBox1);
            reset(textBox2);
            reset(textBox5);
            reset(textBox3);
            reset(textBox6);
            reset(textBox7);
            reset(textBox8);
            reset(textBox9);
        }

        private void reset(TextBox tb)
        {
            tb.Text = null;
            tb.BackColor = Color.White;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            WriteToExcel();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ActiveControl = groupBox1;
            this.Size = new Size(315, 412);
            comboBox1.Text = Convert.ToString(comboBox1.Items[0]);
            textBox4.Text = Properties.Settings.Default.path;
            textBox11.Text = Properties.Settings.Default.password;
            textBox10.Text = Properties.Settings.Default.name;
            Password.test = Properties.Settings.Default.password;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                Form2 form2 = new Form2();
                form2.ShowDialog();
                if (Password.password) this.Size = new Size(315, 525);
                if (!Password.password) checkBox1.Checked = false;
                Password.password = false;
            }
            else
            {
                this.Size = new Size(315, 412);
                Password.test = textBox11.Text;
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            short result;
            if (!String.IsNullOrEmpty(textBox5.Text))
            {
                if (!short.TryParse(textBox5.Text, out result)) textBox5.BackColor = Color.LightCoral;
                else textBox5.BackColor = Color.LightGreen;
            }
            else textBox5.BackColor = Color.White;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            DateTime result;
            if (!String.IsNullOrEmpty(textBox3.Text))
            {
                if (!DateTime.TryParse(textBox3.Text, out result)) textBox3.BackColor = Color.LightCoral;
                else textBox3.BackColor = Color.LightGreen;
            }
            else textBox3.BackColor = Color.White;
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            DateTime result;
            if (!String.IsNullOrEmpty(textBox6.Text))
            {
                if (!DateTime.TryParse(textBox6.Text, out result)) textBox6.BackColor = Color.LightCoral;
                else textBox6.BackColor = Color.LightGreen;
            }
            else textBox6.BackColor = Color.White;
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            short result;
            if (!String.IsNullOrEmpty(textBox7.Text))
            {
                if (!short.TryParse(textBox7.Text, out result)) textBox7.BackColor = Color.LightCoral;
                else textBox7.BackColor = Color.LightGreen;
            }
            else textBox7.BackColor = Color.White;
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            DateTime result;
            if (!String.IsNullOrEmpty(textBox8.Text))
            {
                if (!DateTime.TryParse(textBox8.Text, out result)) textBox8.BackColor = Color.LightCoral;
                else textBox8.BackColor = Color.LightGreen;
            }
            else textBox8.BackColor = Color.White;
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            int result;
            if (!String.IsNullOrEmpty(textBox9.Text))
            {
                if (!int.TryParse(textBox9.Text, out result)) textBox9.BackColor = Color.LightCoral;
                else textBox9.BackColor = Color.LightGreen;
            }
            else textBox9.BackColor = Color.White;
        }

        private void CloseFileExcel()
        {
            ObjExcel.ActiveWorkbook.Save();
            ObjWorkBook.Close();
            ObjExcel.Quit();
            ObjExcel = null;
            ObjWorkBook = null;
            ObjWorkSheet = null;
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process p in procs) p.Kill();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            bool changeColor = true;
            if (String.IsNullOrEmpty(textBox1.Text)) textBox1.BackColor = Color.White;
            for (int i = 0; i < textBox1.TextLength; i++)
            {
                if (!Char.IsLetter(textBox1.Text, i))
                {
                    textBox1.BackColor = Color.LightCoral;
                    changeColor = false;
                }
                if (changeColor) textBox1.BackColor = Color.LightGreen;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            bool changeColor = true;
            if (String.IsNullOrEmpty(textBox2.Text)) textBox2.BackColor = Color.White;
            for (int i = 0; i < textBox2.TextLength; i++)
            {
                if (!Char.IsLetter(textBox2.Text, i))
                {
                    textBox2.BackColor = Color.LightCoral;
                    changeColor = false;
                }
                if (changeColor) textBox2.BackColor = Color.LightGreen;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            reset(textBox1);
            reset(textBox2);
            reset(textBox5);
            reset(textBox3);
            reset(textBox6);
            reset(textBox7);
            reset(textBox8);
            reset(textBox9);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.path = textBox4.Text;
            Properties.Settings.Default.name = textBox10.Text;
            Properties.Settings.Default.password = textBox11.Text;
            Properties.Settings.Default.Save();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                textBox11.PasswordChar = '\0';
                checkBox2.Text = "Скрыть";
            }
            else
            {
                textBox11.PasswordChar = '*';
                checkBox2.Text = "Показать";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            string NameExcel = @"" + textBox4.Text + "\\" + textBox10.Text + ".xlsx";
            try
            {
                ObjWorkBook = ObjExcel.Workbooks.Open(NameExcel,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
                MessageBox.Show("Файл найден и его можно использовать");
                ObjWorkBook.Close();
                ObjExcel.Quit();
                ObjExcel = null;
                ObjWorkBook = null;
                System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (System.Diagnostics.Process p in procs) p.Kill();
            }
            catch
            {
                MessageBox.Show("Неверный путь к файлу", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
    }
}
