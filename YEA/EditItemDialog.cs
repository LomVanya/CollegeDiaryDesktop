using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace YEA
{
    public partial class EditItemDialog : Form
    {
        public string Date { get; set; }
        public string Number { get; set; }
        public string Name { get; set; }
        public string Quantity { get; set; }

        public EditItemDialog(string date, string number, string name, string quantity)
        {
            InitializeComponent();

            Date = date;
            Number = number;
            Name = name;
            Quantity = quantity;

            // Устанавливаем значения текстовых полей
            textBox1.Text = date;
            textBox2.Text = number;
            textBox3.Text = name;
            textBox4.Text = quantity;
        }

      

        private void EditItemDialog_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            // При закрытии формы сохраняем изменения
            Date = textBox1.Text;
            Number = textBox2.Text;
            Name = textBox3.Text;
            Quantity = textBox4.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            Date = textBox1.Text;
            Number = textBox2.Text;
            Name = textBox3.Text;
            Quantity = textBox4.Text;

            Close();
        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value.DayOfWeek == DayOfWeek.Sunday)
            {
                MessageBox.Show("Выберите другую дату, воскресенье недоступно.");
            }
            else
            {
                textBox1.Text = dateTimePicker1.Value.ToString("dd/MM/yyyy");
            }

        }
    }

}
