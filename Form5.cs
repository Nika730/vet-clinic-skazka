using Word = Microsoft.Office.Interop.Word;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;

namespace Pets
{
    public partial class Form5 : Form
    {
        public Form5()
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
                            historys.id,
                            historys.date_historys as 'Дата',
                            historys.description_historys as 'Описание',
                            animals.Nickname as 'Кличка питомца',
                            doctors.FIO as 'Сотрудник'
                        FROM 
                            historys
                        JOIN 
                            animals ON historys.animal_id = animals.id
                        JOIN 
                           doctors ON historys.doctor_id = doctors.id";

                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    dataGridView1.DataSource = table;
                    if (dataGridView1.Columns["Описание"] != null) dataGridView1.Columns["Описание"].Width = 350;
                    if (dataGridView1.Columns["Сотрудник"] != null) dataGridView1.Columns["Сотрудник"].Width = 287;
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
       
       private void Form5_Shown(object sender, EventArgs e)
        {
            LoadPetsData();
            LoadComboBoxes();
            
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
            /////////////////УДАЛЕНИЕ
            ///
        private void button10_Click(object sender, EventArgs e)
        {
            // Проверяем, выбрана ли строка
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Выберите историю для удаления!");
                return;
            }

            // Получаем id выбранной строки
            int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells["id"].Value);

            // Подтверждение удаления
            DialogResult result = MessageBox.Show("Вы уверены, что хотите удалить историю?",
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
                    string query = "DELETE FROM historys WHERE id=@id";
                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@id", id);
                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("История успешно удалена!");
                    LoadPetsData(); // обновляем таблицу
                    ClearFields();  
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
        /////////////////comboBox заполнение
        ///
        private void LoadComboBoxes()
        {
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // Питомцы
                    string queryAnimals = "SELECT id, Nickname FROM animals";
                    MySqlDataAdapter adapterAnimals = new MySqlDataAdapter(queryAnimals, connection);
                    DataTable tableAnimals = new DataTable();
                    adapterAnimals.Fill(tableAnimals);

                    comboBox3.DataSource = tableAnimals;
                    comboBox3.DisplayMember = "Nickname"; // то, что видно
                    comboBox3.ValueMember = "id";         // то, что используется в запросе

                    // Лечение
                    string queryTherapy = "SELECT id, FIO FROM doctors";
                    MySqlDataAdapter adapterTherapy = new MySqlDataAdapter(queryTherapy, connection);
                    DataTable tableTherapy = new DataTable();
                    adapterTherapy.Fill(tableTherapy);

                    comboBox1.DataSource = tableTherapy;
                    comboBox1.DisplayMember = "FIO";
                    comboBox1.ValueMember = "id";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка загрузки ComboBox: " + ex.Message);
                }
            }
        }
        ///
        /////////////////ДОБАВЛЕНИЕ
        ///
        private void button7_Click(object sender, EventArgs e)
        {
            // Проверка на пустые поля
            if (string.IsNullOrWhiteSpace(textBox8.Text) || comboBox1.SelectedValue == null || comboBox3.SelectedValue == null)
            {
                MessageBox.Show("Заполните все поля!");
                return;
            }

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = @"INSERT INTO historys (date_historys, description_historys, animal_id, doctor_id) 
                             VALUES (@date, @description, @animal_id, @doctor_id)";

                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@date", dateTimePicker1.Value);
                        cmd.Parameters.AddWithValue("@description", textBox8.Text);
                        cmd.Parameters.AddWithValue("@animal_id", comboBox3.SelectedValue);
                        cmd.Parameters.AddWithValue("@doctor_id", comboBox1.SelectedValue);

                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("История успешно добавлена!");
                    LoadPetsData(); // обновляем таблицу
                    ClearFields();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка добавления: " + ex.Message);
                }
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                dateTimePicker1.Value = Convert.ToDateTime(row.Cells["Дата"].Value);
                textBox8.Text = row.Cells["Описание"].Value.ToString();

                // comboBox3 — питомец
                comboBox3.Text = row.Cells["Кличка питомца"].Value.ToString();

                // comboBox1 — лечение
                comboBox1.Text = row.Cells["Лечение"].Value.ToString();
            }
        }

        private void графическийРедакторToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Проверка: выбрана ли строка
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Выберите историю для изменения!", "Изменение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Получение ID записи
            int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells["id"].Value);

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    string query = @"UPDATE historys 
                             SET date_historys = @date, 
                                 description_historys = @description, 
                                 animal_id = @animal_id, 
                                 doctor_id = @doctor_id 
                             WHERE id = @id";

                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@date", dateTimePicker1.Value.Date);
                        cmd.Parameters.AddWithValue("@description", textBox8.Text);
                        cmd.Parameters.AddWithValue("@animal_id", comboBox3.SelectedValue);
                        cmd.Parameters.AddWithValue("@doctor_id", comboBox1.SelectedValue);
                        cmd.Parameters.AddWithValue("@id", id);

                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("История успешно изменена!", "Изменение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadPetsData(); // обновление таблицы
                    ClearFields();  // очистка полей, если есть такой метод
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка изменения: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void ClearFields()
        {
            // Очистка текстовых полей
            textBox8.Clear(); // Кличка
            textBox1.Clear(); // Кличка


            // Сброс ComboBox
            comboBox1.SelectedIndex = -1; // Пол
            comboBox3.SelectedIndex = -1; // Владелец

            // Сброс даты (например, на сегодня)
            dateTimePicker1.Value = DateTime.Today;
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

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadPetsData();
            ClearFields();  
            comboBox2.Visible = false;
            button11.Visible = false;
            label7.Visible = false;
           
        }

        private void фильтрацияПоКличкеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox2.Visible = true;
            button11.Visible = true;
            label7.Visible = true;

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string nicknameQuery = "SELECT DISTINCT Nickname FROM animals ORDER BY Nickname";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(nicknameQuery, connection);
                    DataTable table = new DataTable();
                    adapter.Fill(table);

                    comboBox2.DataSource = table;
                    comboBox2.DisplayMember = "Nickname";
                    comboBox2.ValueMember = "Nickname";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка загрузки кличек: " + ex.Message);
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedValue == null)
            {
                MessageBox.Show("Выберите кличку питомца для фильтрации!");
                return;
            }

            string selectedNickname = comboBox2.SelectedValue.ToString();

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = @"
                SELECT 
                    historys.id,
                    historys.date_historys AS 'Дата',
                    historys.description_historys AS 'Описание',
                    animals.Nickname AS 'Кличка питомца',
                    doctors.FIO AS 'Сотрудник'
                FROM historys
                JOIN animals ON historys.animal_id = animals.id
                JOIN doctors ON historys.doctor_id = doctors.id
                WHERE animals.Nickname = @Nickname";

                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    adapter.SelectCommand.Parameters.AddWithValue("@Nickname", selectedNickname);

                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    dataGridView1.DataSource = table;

                    if (dataGridView1.Columns.Contains("id"))
                        dataGridView1.Columns["id"].Visible = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка фильтрации: " + ex.Message);
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

        private void историяБолезниToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedValue == null)
            {
                MessageBox.Show("Выберите питомца для экспорта!");
                return;
            }

            string nickname = comboBox2.SelectedValue.ToString();

            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    connection.Open();
                    string query = @"
                SELECT 
                    historys.date_historys AS 'Дата',
                    historys.description_historys AS 'Описание',
                    animals.Nickname AS 'Кличка питомца',
                    doctors.FIO AS 'Сотрудник'
                FROM historys
                JOIN animals ON historys.animal_id = animals.id
                JOIN doctors ON historys.doctor_id = doctors.id
                WHERE animals.Nickname = @Nickname";

                    MySqlCommand cmd = new MySqlCommand(query, connection);
                    cmd.Parameters.AddWithValue("@Nickname", nickname);

                    MySqlDataReader reader = cmd.ExecuteReader();

                    Word.Application wordApp = new Word.Application();
                    wordApp.Visible = true;
                    Word.Document doc = wordApp.Documents.Add();

                    // Настройка страницы
                    doc.PageSetup.PageWidth = wordApp.CentimetersToPoints(21f);
                    doc.PageSetup.PageHeight = wordApp.CentimetersToPoints(29.7f);
                    doc.PageSetup.TopMargin = wordApp.CentimetersToPoints(2f);
                    doc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(2f);
                    doc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(2.5f);
                    doc.PageSetup.RightMargin = wordApp.CentimetersToPoints(2.5f);

                    // Заголовок
                    Word.Range range = doc.Range(0, 0);
                    range.Text = $"ИСТОРИЯ БОЛЕЗНИ\nПитомец: {nickname}\nДата формирования: {DateTime.Now:dd.MM.yyyy}\n";
                    range.Font.Size = 14;
                    range.Font.Bold = 1;
                    range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    range.InsertParagraphAfter();

                    // Таблица
                    Word.Table table = doc.Tables.Add(range.Paragraphs.Last.Range, 1, 4);
                    table.Range.Font.Size = 11;
                    table.Borders.Enable = 1;

                    string[] headers = { "Дата", "Описание", "Кличка питомца", "Сотрудник" };
                    float[] columnWidths = { 2.5f, 7f, 3f, 4f };

                    for (int i = 0; i < headers.Length; i++)
                    {
                        table.Cell(1, i + 1).Range.Text = headers[i];
                        table.Cell(1, i + 1).Range.Bold = 1;
                        table.Cell(1, i + 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        table.Cell(1, i + 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        table.Columns[i + 1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                        table.Columns[i + 1].PreferredWidth = wordApp.CentimetersToPoints(columnWidths[i]);
                    }

                    int rowIndex = 2;
                    while (reader.Read())
                    {
                        table.Rows.Add();
                        table.Cell(rowIndex, 1).Range.Text = Convert.ToDateTime(reader["Дата"]).ToString("dd.MM.yyyy");
                        table.Cell(rowIndex, 2).Range.Text = reader["Описание"].ToString();
                        table.Cell(rowIndex, 3).Range.Text = reader["Кличка питомца"].ToString();
                        table.Cell(rowIndex, 4).Range.Text = reader["Сотрудник"].ToString();

                        for (int i = 1; i <= 4; i++)
                        {
                            table.Cell(rowIndex, i).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            table.Cell(rowIndex, i).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        }

                        rowIndex++;
                    }

                    reader.Close();

                    // Подпись врача и печать
                    Word.Paragraph signBlock = doc.Paragraphs.Add();
                    signBlock.Range.Text =
                        "\n_________________________\n" +
                        "     Подпись врача\n\n" +
                        "М.П.";
                    signBlock.Range.Font.Size = 12;
                    signBlock.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка экспорта: " + ex.Message);
            }
        }

    }

}













    

