using System;
using System.Linq;
using System.Windows.Forms;

namespace Tab_export
{
    public partial class bad_tab_form : Form
    {
        public bad_tab_form()
        {
            InitializeComponent();
            textBox1.Text = Properties.Settings.Default.tab_prec.ToString();
            textBox2.Text = Properties.Settings.Default.cell_prec.ToString();
            checkBox1.Checked = Properties.Settings.Default.cell_po_centru;
            checkBox2.Checked = Properties.Settings.Default.cell_format;
        }

        /// <summary>
        /// добавляет 0 к содержимому текстбокса если в нем нет чисел
        /// используется при событии потери фокуса что бы в текстбоксе точно было число которое можно парсить      
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void TbLostFocus(object sender, EventArgs e)
        {
            TextBox tTB = sender as TextBox;
            foreach (Char c in tTB.Text) if (Char.IsDigit(c)) return;
            tTB.Text += "0";
        }
        /// <summary>
        /// это событие позволяет возможность вводить десятичные положительные и отрицательные числа
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void TbKeyDoubleMinus(object sender, KeyPressEventArgs e)
        {
            TextBox tTB = sender as TextBox;
            if (e.KeyChar == ',') e.KeyChar = '.';
            if (e.KeyChar == '-')
            {
                int pos = tTB.SelectionStart;
                if (tTB.Text.Contains('-'))
                {
                    tTB.Text = tTB.Text.Substring(1);
                    tTB.SelectionStart = pos - 1;
                }
                else
                {
                    tTB.Text = '-' + tTB.Text;
                    tTB.SelectionStart = pos + 1;
                }
                e.Handled = true;
                return;
            }
            if (e.KeyChar == '.')
            {
                if (tTB.Text.Contains('.') | (tTB.SelectionStart == 0 & tTB.Text.Contains('-'))) e.Handled = true;
                return;
            }
            if (e.KeyChar == 8)
            {
                if (tTB.SelectionLength > 0)
                {
                    int pos = tTB.SelectionStart;
                    tTB.Text = tTB.Text.Substring(0, tTB.SelectionStart) + tTB.Text.Substring(tTB.SelectionStart + tTB.SelectionLength);
                    tTB.SelectionStart = pos;
                    e.Handled = true;
                }
                return;
            }
            if (Char.IsDigit(e.KeyChar))
            {
                if (tTB.Text.Contains('-') & tTB.SelectionStart == 0) e.Handled = true;
                return;
            }
            e.Handled = true;
        }
        /// <summary>
        /// это событие позволяет вводить десятичные положительные числа
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void TbKeyDouble(object sender, KeyPressEventArgs e)
        {
            TextBox tTB = sender as TextBox;
            if (e.KeyChar == ',') e.KeyChar = '.';
            if (e.KeyChar == '.')
            {
                if (tTB.Text.Contains('.') | (tTB.SelectionStart == 0 & tTB.Text.Contains('-'))) e.Handled = true;
                return;
            }
            if (e.KeyChar == 8)
            {
                if (tTB.SelectionLength > 0)
                {
                    int pos = tTB.SelectionStart;
                    tTB.Text = tTB.Text.Substring(0, tTB.SelectionStart) + tTB.Text.Substring(tTB.SelectionStart + tTB.SelectionLength);
                    tTB.SelectionStart = pos;
                    e.Handled = true;
                }
                return;
            }
            if (Char.IsDigit(e.KeyChar))
            {
                if (tTB.Text.Contains('-') & tTB.SelectionStart == 0) e.Handled = true;
                return;
            }
            e.Handled = true;
        }
        /// <summary>
        /// это событие позволяет вводить целые положительные и отрицательные числа
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void TbKeyIntegerMinus(object sender, KeyPressEventArgs e)
        {
            TextBox tTB = sender as TextBox;
            if (e.KeyChar == '-')
            {
                int pos = tTB.SelectionStart;
                if (tTB.Text.Contains('-'))
                {
                    tTB.Text = tTB.Text.Substring(1);
                    tTB.SelectionStart = pos - 1;
                }
                else
                {
                    tTB.Text = '-' + tTB.Text;
                    tTB.SelectionStart = pos + 1;
                }
                e.Handled = true;
                return;
            }
            if (e.KeyChar == 8)
            {
                if (tTB.SelectionLength > 0)
                {
                    int pos = tTB.SelectionStart;
                    tTB.Text = tTB.Text.Substring(0, tTB.SelectionStart) + tTB.Text.Substring(tTB.SelectionStart + tTB.SelectionLength);
                    tTB.SelectionStart = pos;
                    e.Handled = true;
                }
                return;
            }
            if (Char.IsDigit(e.KeyChar))
            {
                if (tTB.Text.Contains('-') & tTB.SelectionStart == 0) e.Handled = true;
                return;
            }
            e.Handled = true;
        }
        /// <summary>
        /// это событие позволяет вводить целые положительные числа
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void TbKeyInteger(object sender, KeyPressEventArgs e)
        {
            TextBox tTB = sender as TextBox;
            if (e.KeyChar == 8)
            {
                if (tTB.SelectionLength > 0)
                {
                    int pos = tTB.SelectionStart;
                    tTB.Text = tTB.Text.Substring(0, tTB.SelectionStart) + tTB.Text.Substring(tTB.SelectionStart + tTB.SelectionLength);
                    tTB.SelectionStart = pos;
                    e.Handled = true;
                }
                return;
            }
            if (Char.IsDigit(e.KeyChar))
            {
                if (tTB.Text.Contains('-') & tTB.SelectionStart == 0) e.Handled = true;
                return;
            }
            e.Handled = true;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double result = 0;
            if (Double.TryParse(textBox1.Text, out result))
            {
                Properties.Settings.Default.tab_prec = result;
            }
            else
            {
                MessageBox.Show("Точность таблицы задана некорректно");
                return;
            }
            if (Double.TryParse(textBox2.Text, out result))
            {
                Properties.Settings.Default.cell_prec = result;
            }
            else
            {
                MessageBox.Show("Точность ячеек задана некоррекно");
                return;
            }
            Properties.Settings.Default.cell_po_centru = checkBox1.Checked;
            Properties.Settings.Default.cell_format = checkBox2.Checked;
            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}
