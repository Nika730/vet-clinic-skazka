using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Mysqlx.Notice.Warning.Types;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Pets
{
    public partial class Form7 : Form
    {
        public Form7()
        {
            InitializeComponent();
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
                    a.id AS 'ID', 
                    a.FIO AS 'ФИО', 
                    a.phone AS 'Телефон', 
                    a.email AS 'Почта' 
                    FROM owners a";

                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    dataGridView1.DataSource = table;
                    if (dataGridView1.Columns["ФИО"] != null) dataGridView1.Columns["ФИО"].Width = 350;
                    if (dataGridView1.Columns["Почта"] != null) dataGridView1.Columns["Почта"].Width = 250;
                    if (dataGridView1.Columns["Телефон"] != null) dataGridView1.Columns["Телефон"].Width = 220;
                    // Настройка стиля заголовков
                    dataGridView1.EnableHeadersVisualStyles = false; // отключаем системный стиль
                    dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.LightBlue; // фон
                    dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.DarkBlue;  // текст
                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold); // шрифт
                    dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // выравнивание

                    if (dataGridView1.Columns.Contains("id"))
                        dataGridView1.Columns["id"].Visible = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка подключения: " + ex.Message);
                }
            }
        }
      
        private void Form7_Shown(object sender, EventArgs e)
        {
            LoadPetsData();
            
        }

        private void ClearFields()
        {
            textBox2.Clear(); // Кличка
            textBox3.Clear(); // Порода
            textBox4.Clear();
            textBox1.Clear();// Вид
        }
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
            this.Hide();
            this.Close();
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.Show();
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

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            Application.Exit();
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


        ///
        /////////////////ДОБАВЛЕНИЕ
        ///
        private void button7_Click(object sender, EventArgs e)
        {
            if (!textBox2.MaskFull)
            {
                MessageBox.Show("Введите полный номер телефона!");
                return;
            }
            string fio = textBox3.Text.Trim();
            string specially = textBox4.Text.Trim();
            string phone = textBox2.Text.Trim(); // maskedTextBox

            // Проверка на заполненность
            if (string.IsNullOrWhiteSpace(fio) || string.IsNullOrWhiteSpace(specially) || string.IsNullOrWhiteSpace(phone))
            {
                MessageBox.Show("Пожалуйста, заполните все поля.");
                return;
            }

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = "INSERT INTO owners (FIO, phone, email) VALUES (@fio, @phone, @spec)";
                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@fio", fio);
                        cmd.Parameters.AddWithValue("@spec", specially);
                        cmd.Parameters.AddWithValue("@phone", phone);
                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Владелец успешно добавлен!");
                    LoadPetsData(); // обновить таблицу
                    ClearFields();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при добавлении: " + ex.Message);
                }
            }
        }
        ///
        /////////////////УДАЛЕНИЕ
        ///
        private void button10_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Выберите владельца для удаления!");
                return;
            }
            ClearFields();
            // Получаем id выбранной строки
            int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells["id"].Value);

            // Подтверждение удаления
            DialogResult result = MessageBox.Show("Вы уверены, что хотите удалить владельца?",
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
                    string query = "DELETE FROM owners WHERE id=@id";
                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@id", id);
                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Владелец успешно удален!");
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

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadPetsData();
            ClearFields();
        }

        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Проверяем, выбрана ли строка
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Выберите сотрудника для изменения!");
                return;
            }
            if (string.IsNullOrWhiteSpace(textBox2.Text) ||
                string.IsNullOrWhiteSpace(textBox3.Text) ||
                string.IsNullOrWhiteSpace(textBox4.Text))
            {
                MessageBox.Show("Пожалуйста, заполните все поля перед добавлением питомца!");
                return;
            }
            // Получаем id выбранной строки
            int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells["id"].Value);

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    string query = @"UPDATE owners 
                             SET FIO=@fio, 
                                 phone=@phone,
                                    email=@spec                                   
                             WHERE id=@id";

                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@fio", textBox3.Text.Trim());
                        cmd.Parameters.AddWithValue("@spec", textBox4.Text.Trim());
                        cmd.Parameters.AddWithValue("@phone", textBox2.Text.Trim());
                        cmd.Parameters.AddWithValue("@id", id);

                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Сотрудник успешно обновлён!");
                    LoadPetsData(); // обновляем таблицу
                    ClearFields();  // если есть метод очистки
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка изменения: " + ex.Message);
                }
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                textBox3.Text = row.Cells["ФИО"].Value.ToString();
                textBox4.Text = row.Cells["Почта"].Value.ToString();
                textBox2.Text = row.Cells["Телефон"].Value.ToString();
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

        private void button9_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
            this.Hide();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.Show();
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

        
    }
}
