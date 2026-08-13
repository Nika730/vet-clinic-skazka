using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Mysqlx.Notice.Warning.Types;

namespace Pets
{ 
    public partial class Form3 : Form
    {


        public Form3()
        { 
            InitializeComponent();
            textBox1.TextChanged += textBox1_TextChanged;
            textBox1.KeyDown += textBox1_KeyDown;
        }

        private static string connectionString = "server=localhost;user=root;password=admin;database=pet;";
        private void LoadPetsData()
        {

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = @"
                    SELECT 
                    services.id,
                    services.name_services as 'Название',
                    services.description_services as 'Описание',
                    services.price as 'Цена'
                    FROM services";

                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    dataGridView1.DataSource = table;
                    
                    if (dataGridView1.Columns.Contains("id"))
                        dataGridView1.Columns["id"].Visible = false;
                    // Настройка стиля заголовков
                    dataGridView1.EnableHeadersVisualStyles = false; // отключаем системный стиль
                    dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.LightBlue; // фон
                    dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.DarkBlue;  // текст
                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold); // шрифт
                    dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // выравнивание


                    if (dataGridView1.Columns["Название"] != null) dataGridView1.Columns["Название"].Width = 250;
                    if (dataGridView1.Columns["Описание"] != null) dataGridView1.Columns["Описание"].Width = 470;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка подключения: " + ex.Message);
                }
            }
        }
        

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form3_Shown(object sender, EventArgs e)
        {
            LoadPetsData();
           
        }
        ///
        /////////////////ДОБАВЛЕНИЕ
        ///
        private void button7_Click(object sender, EventArgs e)
        {
            // Проверка на пустые поля
            if (string.IsNullOrWhiteSpace(textBox8.Text) ||
                string.IsNullOrWhiteSpace(textBox4.Text) ||
                string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("Все поля должны быть заполнены!");
                return;
            }

            // Проверка цены через регулярное выражение
            string priceInput = textBox4.Text.Trim();
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"^\d{1,3}(\,\d{1,2})?$");

            if (!regex.IsMatch(priceInput))
            {
                MessageBox.Show("Цена должна быть в формате: до 3 цифр до запятой и до 2 после (например: 123,45)");
                return;
            }

            // Преобразуем цену в decimal (заменяем запятую на точку для MySQL)
            string normalizedPrice = priceInput.Replace(',', '.');

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = "INSERT INTO services (name_services, description_services, price) VALUES (@name, @desc, @price)";
                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@name", textBox2.Text.Trim());
                        cmd.Parameters.AddWithValue("@desc", textBox8.Text.Trim());
                        cmd.Parameters.AddWithValue("@price", normalizedPrice);

                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Услуга успешно добавлена!");
                    LoadPetsData(); // обновляем таблицу

                    // Очистка полей после добавления
                    textBox2.Clear();   // имя
                    textBox8.Clear();   // описание
                    textBox4.Clear();   // цена
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка добавления: " + ex.Message);
                }
            }
        }
        ///
        /////////////////ОГРАНИЕЧЕНИЕ ДЛЯ ЦЕНЫ
        ///
        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Разрешаем цифры, запятую и Backspace
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != ',' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
                return;
            }

            // Если нажата клавиша Backspace — разрешаем без проверки
            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
                return;
            }

            // Запрещаем запятую как первый символ
            if (e.KeyChar == ',' && tb.SelectionStart == 0)
            {
                e.Handled = true;
                return;
            }

            // Разрешаем только одну запятую
            if (e.KeyChar == ',' && tb.Text.Contains(","))
            {
                e.Handled = true;
                return;
            }

            // Смоделируем будущий текст после ввода символа
            string current = tb.Text;
            int selStart = tb.SelectionStart;
            int selLength = tb.SelectionLength;
            string next = current.Substring(0, selStart) + e.KeyChar + current.Substring(selStart + selLength);

            // Проверка формата: до 3 цифр до запятой и до 2 после
            Regex rx = new Regex(@"^\d{0,3}(,\d{0,2})?$");
            if (!rx.IsMatch(next))
            {
                e.Handled = true;
                return;
            }
        }
        ///
        /////////////////УДАЛЕНИЕ
        ///
        private void button10_Click(object sender, EventArgs e)
        {
            // Проверяем, выбрана ли строка
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Выберите услугу для удаления!");
                return;
            }

            // Получаем id выбранной строки
            int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells["id"].Value);

            // Подтверждение удаления
            DialogResult result = MessageBox.Show("Вы уверены, что хотите удалить услугу?",
                                                  "Подтверждение",
                                                  MessageBoxButtons.YesNo,
                                                  MessageBoxIcon.Question);

            if (result == DialogResult.No)
                return;

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = "DELETE FROM services WHERE id=@id";
                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@id", id);
                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Услуга успешно удалена!");
                    LoadPetsData(); // обновляем таблицу
                }
                catch (MySqlException ex)
                {
                    if (ex.Number == 1451) // Ошибка внешнего ключа
                    {
                        MessageBox.Show("Невозможно удалить запись, так как она используется в другой таблице.",
                                        "Удаление невозможно",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Warning);
                    }
                    else
                    {
                        MessageBox.Show("Ошибка при удалении: " + ex.Message,
                                        "Ошибка",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка удаления: " + ex.Message);
                }
            }
        }
        ///
        /////////////////ФИЛЬТРАЦИЯ
        ///
        private void графическийРедакторToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label4.Visible = true; // "От"
            label5.Visible = true; // "До"
            textBox3.Visible = true; // поле "от"
            textBox5.Visible = true; // поле "до"
            button11.Visible = true; // сама кнопка фильтрации
        }

        private void button11_Click(object sender, EventArgs e)
        {
            // Проверка на пустые поля
            if (string.IsNullOrWhiteSpace(textBox3.Text) || string.IsNullOrWhiteSpace(textBox5.Text))
            {
                MessageBox.Show("Введите диапазон цен!");
                return;
            }

            // Преобразуем значения (заменяем запятую на точку)
            string fromPrice = textBox3.Text.Trim().Replace(',', '.');
            string toPrice = textBox5.Text.Trim().Replace(',', '.');

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = @"
                SELECT 
                    services.id,
                    services.name_services AS 'Название',
                    services.description_services AS 'Описание',
                    services.price AS 'Цена'
                FROM services
                WHERE price BETWEEN @from AND @to";

                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    adapter.SelectCommand.Parameters.AddWithValue("@from", fromPrice);
                    adapter.SelectCommand.Parameters.AddWithValue("@to", toPrice);

                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    dataGridView1.DataSource = table;

                    if (dataGridView1.Columns.Contains("id"))
                        dataGridView1.Columns["id"].Visible = false;

                    if (dataGridView1.Columns["Название"] != null) dataGridView1.Columns["Название"].Width = 250;
                    if (dataGridView1.Columns["Описание"] != null) dataGridView1.Columns["Описание"].Width = 470;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка фильтрации: " + ex.Message);
                }
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Разрешаем цифры, запятую и Backspace
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != ',' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
                return;
            }

            // Если нажата клавиша Backspace — разрешаем без проверки
            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
                return;
            }

            // Запрещаем запятую как первый символ
            if (e.KeyChar == ',' && tb.SelectionStart == 0)
            {
                e.Handled = true;
                return;
            }

            // Разрешаем только одну запятую
            if (e.KeyChar == ',' && tb.Text.Contains(","))
            {
                e.Handled = true;
                return;
            }

            // Смоделируем будущий текст после ввода символа
            string current = tb.Text;
            int selStart = tb.SelectionStart;
            int selLength = tb.SelectionLength;
            string next = current.Substring(0, selStart) + e.KeyChar + current.Substring(selStart + selLength);

            // Проверка формата: до 3 цифр до запятой и до 2 после
            Regex rx = new Regex(@"^\d{0,3}(,\d{0,2})?$");
            if (!rx.IsMatch(next))
            {
                e.Handled = true;
                return;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox tb = sender as TextBox;

            // Разрешаем цифры, запятую и Backspace
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != ',' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
                return;
            }

            // Если нажата клавиша Backspace — разрешаем без проверки
            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
                return;
            }

            // Запрещаем запятую как первый символ
            if (e.KeyChar == ',' && tb.SelectionStart == 0)
            {
                e.Handled = true;
                return;
            }

            // Разрешаем только одну запятую
            if (e.KeyChar == ',' && tb.Text.Contains(","))
            {
                e.Handled = true;
                return;
            }

            // Смоделируем будущий текст после ввода символа
            string current = tb.Text;
            int selStart = tb.SelectionStart;
            int selLength = tb.SelectionLength;
            string next = current.Substring(0, selStart) + e.KeyChar + current.Substring(selStart + selLength);

            // Проверка формата: до 3 цифр до запятой и до 2 после
            Regex rx = new Regex(@"^\d{0,3}(,\d{0,2})?$");
            if (!rx.IsMatch(next))
            {
                e.Handled = true;
                return;
            }
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadPetsData();
            textBox2.Clear();   // имя
            textBox8.Clear();   // описание
            textBox4.Clear();
            textBox1.Clear();
            label4.Visible = false;
            label5.Visible = false;
            textBox3.Visible = false;
            textBox5.Visible = false;
            button11.Visible = false;

        }

        ///
        /////////////////ПОИСК
        ///
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string searchText = textBox1.Text.Trim().ToLower();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.DefaultCellStyle.BackColor = Color.White; // сброс цвета

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && cell.Value.ToString().ToLower().Contains(searchText))
                    {
                        row.DefaultCellStyle.BackColor = Color.LightYellow; // подсветка
                        break;
                    }
                }
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string searchText = textBox1.Text.Trim().ToLower();

                if (dataGridView1.DataSource is DataTable table)
                {
                    // Добавляем временный столбец Match, если его нет
                    if (!table.Columns.Contains("Match"))
                        table.Columns.Add("Match", typeof(int));

                    foreach (DataRow row in table.Rows)
                    {
                        bool matchFound = row.ItemArray.Any(val =>
                            val != null && val.ToString().ToLower().Contains(searchText));

                        row["Match"] = matchFound ? 0 : 1; // 0 — совпадение, 1 — нет
                    }

                    // Сортируем по Match
                    DataView view = table.DefaultView;
                    view.Sort = "Match ASC";
                    dataGridView1.DataSource = view;
                    if (dataGridView1.Columns.Contains("Match"))
                        dataGridView1.Columns["Match"].Visible = false;
                    // Очистка строки поиска
                    textBox1.Clear();
                    // Снимаем выделение
                    dataGridView1.ClearSelection();

                    // Убираем подсветку у всех строк
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.DefaultCellStyle.BackColor = Color.White;
                    }
                                       
                    // Предотвращаем "звон" Enter
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            }
        }

        
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
            this.Hide();
            this.Close();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Form11 form11 = new Form11();
            form11.Show();
            this.Hide();
            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form6 form6 = new Form6();
            form6.Show();
            this.Hide();
            this.Close();
        }

      

        private void button4_Click(object sender, EventArgs e)
        {
            Form5 form5 = new Form5();
            form5.Show();
            this.Hide();
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form4 form4 = new Form4();
            form4.Show();
            this.Hide();
            this.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form7 form7 = new Form7();
            form7.Show();
            this.Hide();
            this.Close();
        }

        private void записьНаПриемToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form11 form11 = new Form11();
            form11.Show();
            this.Hide();
            this.Close();
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            Form6 form6 = new Form6();
            form6.Show();
            this.Hide();
            this.Close();
        }

      

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            Form5 form5 = new Form5();
            form5.Show();
            this.Hide();
            this.Close();
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            Form4 form4 = new Form4();
            form4.Show();
            this.Hide();
            this.Close();
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            Form7 form7 = new Form7();
            form7.Show();
            this.Hide();
            this.Close();
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            Form8 form8 = new Form8();
            form8.Show();
            this.Hide();
            this.Close();
        }

        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {
            Form9 form9 = new Form9();
            form9.Show();
            this.Hide();
            this.Close();
        }

        private void exelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Columns.ColumnWidth = 15;

            Microsoft.Office.Interop.Excel.Worksheet wsh = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveSheet;

            // Заголовок таблицы
            ExcelApp.Cells[1, 1] = "УСЛУГИ ВЕТКЛИНИКИ";
            Microsoft.Office.Interop.Excel.Range titleRange = wsh.Range[wsh.Cells[1, 1], wsh.Cells[1, dataGridView1.ColumnCount - 1]];
            titleRange.Merge();
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;
            titleRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            titleRange.Interior.Color = System.Drawing.Color.LightBlue;
            titleRange.Font.Color = System.Drawing.Color.DarkBlue;

            // Заголовки столбцов (без ID)
            ExcelApp.Cells[2, 1] = "Название";
            ExcelApp.Cells[2, 2] = "Описание";
            ExcelApp.Cells[2, 3] = "Цена";

            // Форматирование заголовков
            Microsoft.Office.Interop.Excel.Range headerRange = wsh.Range[wsh.Cells[2, 1], wsh.Cells[2, 3]];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = System.Drawing.Color.LightGray;
            headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            // Данные (пропускаем ID = колонка 0)
            for (int i = 1; i < dataGridView1.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView1.RowCount - 1; j++)
                {
                    ExcelApp.Cells[j + 3, i] = dataGridView1[i, j].Value?.ToString();
                }

                // Заливка для каждого столбца
                Microsoft.Office.Interop.Excel.Range colRange = wsh.Range[wsh.Cells[3, i], wsh.Cells[dataGridView1.RowCount + 2, i]];
                if (i == 1) colRange.Interior.Color = System.Drawing.Color.LightYellow;
                else if (i == 2) colRange.Interior.Color = System.Drawing.Color.LightCyan;
                else if (i == 3) colRange.Interior.Color = System.Drawing.Color.LightGreen;
            }

            // Авторазмер и границы
            Microsoft.Office.Interop.Excel.Range eRange = wsh.UsedRange;
            eRange.EntireRow.AutoFit();
            eRange.EntireColumn.AutoFit();

            Microsoft.Office.Interop.Excel.Borders border = eRange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // Подпись внизу
            ExcelApp.Cells[dataGridView1.RowCount + 5, dataGridView1.ColumnCount - 2] = "подпись";

            ExcelApp.Visible = true;
        }

        private void wordToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
            this.Hide();
            this.Close();
        }
    }

}









