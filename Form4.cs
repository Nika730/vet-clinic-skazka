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
using Word = Microsoft.Office.Interop.Word;

namespace Pets
{
    public partial class Form4 : Form
    {
        public Form4()
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
                    animals.id,
                    animals.Species as 'Вид',
                    animals.Nickname as 'Кличка',
                    animals.Breed as 'Порода',
                    animals.Gender as 'Пол',
                    animals.DR as 'Дата рождения',
                    owners.FIO as 'Владелец'
                     FROM 
                    animals
                      JOIN 
                    owners ON animals.Owner_ID = owners.id";

                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    dataGridView1.DataSource = table;
                    if (dataGridView1.Columns["Владелец"] != null) dataGridView1.Columns["Владелец"].Width = 230;
                    if (dataGridView1.Columns["Порода"] != null) dataGridView1.Columns["Порода"].Width = 165;
                    if (dataGridView1.Columns["Кличка"] != null) dataGridView1.Columns["Кличка"].Width = 120;
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
        private void Form4_Shown_1(object sender, EventArgs e)
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

   

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            Form5 form5 = new Form5();
            form5.Show();
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
        /////////////////comboBox заполнение
        ///
        private void LoadComboBoxes()
        {
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // Загружаем владельцев
                    string ownersQuery = "SELECT id, FIO FROM owners";
                    MySqlDataAdapter ownersAdapter = new MySqlDataAdapter(ownersQuery, connection);
                    DataTable ownersTable = new DataTable();
                    ownersAdapter.Fill(ownersTable);

                    comboBox3.DataSource = ownersTable;
                    comboBox3.DisplayMember = "FIO";   // отображаем ФИО
                    comboBox3.ValueMember = "id";      // значение — id владельца

                    // Загружаем пол (можно хранить в отдельной таблице или просто задать вручную)
                    string genderQuery = "SELECT DISTINCT Gender FROM animals";
                    MySqlDataAdapter genderAdapter = new MySqlDataAdapter(genderQuery, connection);
                    DataTable genderTable = new DataTable();
                    genderAdapter.Fill(genderTable);

                    comboBox1.DataSource = genderTable;
                    comboBox1.DisplayMember = "Gender";
                    comboBox1.ValueMember = "Gender";
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
            if (string.IsNullOrWhiteSpace(textBox2.Text) ||
            string.IsNullOrWhiteSpace(textBox3.Text) ||
            string.IsNullOrWhiteSpace(textBox4.Text) ||
            comboBox1.SelectedItem == null ||
            comboBox3.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, заполните все поля перед добавлением питомца!");
                return;
            }
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    string query = @"INSERT INTO animals 
                            (Species, Nickname, Breed, Gender, DR, Owner_ID) 
                            VALUES (@Species, @Nickname, @Breed, @Gender, @DR, @Owner_ID)";

                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@Species", textBox4.Text);
                        cmd.Parameters.AddWithValue("@Nickname", textBox2.Text);
                        cmd.Parameters.AddWithValue("@Breed", textBox3.Text);
                        cmd.Parameters.AddWithValue("@Gender", comboBox1.SelectedValue.ToString());
                        cmd.Parameters.AddWithValue("@DR", dateTimePicker1.Value.Date); // ✅ теперь дата берётся из DateTimePicker
                        cmd.Parameters.AddWithValue("@Owner_ID", comboBox3.SelectedValue);

                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Питомец успешно добавлен!");
                    LoadPetsData(); // обновляем таблицу
                    ClearFields();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка добавления: " + ex.Message);
                }
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
                MessageBox.Show("Выберите питомца для удаления!");
                return;
            }
            ClearFields();
            // Получаем id выбранной строки
            int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells["id"].Value);

            // Подтверждение удаления
            DialogResult result = MessageBox.Show("Вы уверены, что хотите удалить питомца?",
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
                    string query = "DELETE FROM animals WHERE id=@id";
                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@id", id);
                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Питомец успешно удален!");
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
        /////////////////ФИЛЬТРАЦИЯ ПО ПОЛУ
        ///
        private void графическийРедакторToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox2.Visible = true;
            button11.Visible = true;
            label7.Visible = true;

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string genderQuery = "SELECT DISTINCT Gender FROM animals";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(genderQuery, connection);
                    DataTable table = new DataTable();
                    adapter.Fill(table);

                    comboBox2.DataSource = table;
                    comboBox2.DisplayMember = "Gender";
                    comboBox2.ValueMember = "Gender";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка загрузки полов: " + ex.Message);
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedValue == null)
            {
                MessageBox.Show("Выберите пол для фильтрации!");
                return;
            }

            string selectedGender = comboBox2.SelectedValue.ToString();

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = @"
                SELECT 
                    animals.id,
                    animals.Species as 'Вид',
                    animals.Nickname as 'Кличка',
                    animals.Breed as 'Порода',
                    animals.Gender as 'Пол',
                    animals.DR as 'Дата рождения',
                    owners.FIO as 'Владелец'
                FROM animals
                JOIN owners ON animals.Owner_ID = owners.id
                WHERE animals.Gender = @Gender";

                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    adapter.SelectCommand.Parameters.AddWithValue("@Gender", selectedGender);

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

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadPetsData();
            ClearFields();  
            comboBox2.Visible = false;
            button11.Visible = false;
            label7.Visible = false;

        }
        ///
        /////////////////ПОИСК
        ///
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

        private void button6_Click(object sender, EventArgs e)
        {
            Form7 form7 = new Form7();
            form7.Show();
            this.Hide();
            this.Close();
        }

        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Проверяем, выбрана ли строка
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Выберите питомца для изменения!");
                return;
            }
            if (string.IsNullOrWhiteSpace(textBox2.Text) ||
            string.IsNullOrWhiteSpace(textBox3.Text) ||
            string.IsNullOrWhiteSpace(textBox4.Text) ||
            comboBox1.SelectedItem == null ||
            comboBox3.SelectedItem == null)
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

                    string query = @"UPDATE animals 
                             SET Species=@Species, 
                                 Nickname=@Nickname, 
                                 Breed=@Breed, 
                                 Gender=@Gender, 
                                 DR=@DR, 
                                 Owner_ID=@Owner_ID
                             WHERE id=@id";

                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        // Берём значения из текстбоксов и комбобоксов
                        cmd.Parameters.AddWithValue("@Species", textBox4.Text);
                        cmd.Parameters.AddWithValue("@Nickname", textBox2.Text);
                        cmd.Parameters.AddWithValue("@Breed", textBox3.Text);
                        cmd.Parameters.AddWithValue("@Gender", comboBox1.SelectedValue.ToString());
                        cmd.Parameters.AddWithValue("@DR", dateTimePicker1.Value.Date);
                        cmd.Parameters.AddWithValue("@Owner_ID", comboBox3.SelectedValue);
                        cmd.Parameters.AddWithValue("@id", id);

                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Запись успешно изменена!");
                    LoadPetsData(); // обновляем таблицу
                    ClearFields();
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

                textBox4.Text = row.Cells["Вид"].Value.ToString();
                textBox2.Text = row.Cells["Кличка"].Value.ToString();
                textBox3.Text = row.Cells["Порода"].Value.ToString();
                comboBox1.SelectedValue = row.Cells["Пол"].Value.ToString();
                dateTimePicker1.Value = Convert.ToDateTime(row.Cells["Дата рождения"].Value);
                comboBox3.Text = row.Cells["Владелец"].Value.ToString();
            }
        }

        private void ClearFields()
        {
            // Очистка текстовых полей
            textBox2.Clear(); // Кличка
            textBox3.Clear(); // Порода
            textBox4.Clear(); // Вид
            textBox1.Clear(); // Вид


            // Сброс ComboBox
            comboBox1.SelectedIndex = -1; // Пол
            comboBox3.SelectedIndex = -1; // Владелец

            // Сброс даты (например, на сегодня)
            dateTimePicker1.Value = DateTime.Today;
        }

        private void ветеринарныйПаспортToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите питомца в таблице!");
                return;
            }

            var row = dataGridView1.SelectedRows[0];

            // Обработка даты
            
            string species = row.Cells["Вид"].Value?.ToString();
            string nickname = row.Cells["Кличка"].Value?.ToString();
            string breed = row.Cells["Порода"].Value?.ToString();
            string gender = row.Cells["Пол"].Value?.ToString();
            string rawDate = row.Cells["Дата рождения"].Value?.ToString();
            DateTime parsedDate;
            string date = DateTime.TryParse(rawDate, out parsedDate)
                ? parsedDate.ToString("dd.MM.yyyy")
                : rawDate;
            string owner = row.Cells["Владелец"].Value?.ToString();

            try
            {
                Word.Application wordApp = new Word.Application();
                wordApp.Visible = true;
                Word.Document doc = wordApp.Documents.Add();

                // Настройка страницы — формат A4
                doc.PageSetup.PageWidth = wordApp.CentimetersToPoints(21f);
                doc.PageSetup.PageHeight = wordApp.CentimetersToPoints(29.7f);
                doc.PageSetup.TopMargin = wordApp.CentimetersToPoints(2f);
                doc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(2f);
                doc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(2.5f);
                doc.PageSetup.RightMargin = wordApp.CentimetersToPoints(2.5f);

                // Заголовок
                Word.Paragraph title = doc.Paragraphs.Add();
                title.Range.Text = "ВЕТЕРИНАРНЫЙ ПАСПОРТ\nКлиника «Сказка»";
                title.Range.Font.Size = 16;
                title.Range.Font.Bold = 1;
                title.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                title.Range.InsertParagraphAfter();

                // Дата выдачи
                Word.Paragraph dateIssued = doc.Paragraphs.Add();
                dateIssued.Range.Text = $"Дата выдачи: {DateTime.Now:dd.MM.yyyy}";
                dateIssued.Range.Font.Size = 12;
                dateIssued.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                dateIssued.Range.InsertParagraphAfter();

                // Таблица с данными питомца
                Word.Table petTable = doc.Tables.Add(doc.Paragraphs.Add().Range, 6, 2);
                petTable.Range.Font.Size = 12;
                petTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowLeft;
                petTable.Columns[1].Width = wordApp.CentimetersToPoints(6f);
                petTable.Columns[2].Width = wordApp.CentimetersToPoints(10f);

                petTable.Cell(1, 1).Range.Text = "Вид:";
                petTable.Cell(1, 2).Range.Text = species;
                petTable.Cell(2, 1).Range.Text = "Кличка:";
                petTable.Cell(2, 2).Range.Text = nickname;
                petTable.Cell(3, 1).Range.Text = "Порода:";
                petTable.Cell(3, 2).Range.Text = breed;
                petTable.Cell(4, 1).Range.Text = "Пол:";
                petTable.Cell(4, 2).Range.Text = gender;
                petTable.Cell(5, 1).Range.Text = "Дата рождения:";
                petTable.Cell(5, 2).Range.Text = date;
                petTable.Cell(6, 1).Range.Text = "Владелец:";
                petTable.Cell(6, 2).Range.Text = owner;

                // Заголовок таблицы прививок
                Word.Paragraph vaccTitle = doc.Paragraphs.Add();
                vaccTitle.Range.Text = "\nПрививки и процедуры:";
                vaccTitle.Range.Font.Size = 12;
                vaccTitle.Range.Font.Bold = 1;
                vaccTitle.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                vaccTitle.Range.InsertParagraphAfter();

                // Таблица прививок
                Word.Table vaccineTable = doc.Tables.Add(doc.Paragraphs.Add().Range, 6, 3);
                vaccineTable.Range.Font.Size = 11;
                vaccineTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                vaccineTable.Columns[1].Width = wordApp.CentimetersToPoints(6f);
                vaccineTable.Columns[2].Width = wordApp.CentimetersToPoints(6f);
                vaccineTable.Columns[3].Width = wordApp.CentimetersToPoints(6f);

                vaccineTable.Cell(1, 1).Range.Text = "Вакцина / процедура";
                vaccineTable.Cell(1, 2).Range.Text = "Дата";
                vaccineTable.Cell(1, 3).Range.Text = "Отметка врача";

                for (int i = 2; i <= 6; i++)
                {
                    vaccineTable.Cell(i, 1).Range.Text = "";
                    vaccineTable.Cell(i, 2).Range.Text = "";
                    vaccineTable.Cell(i, 3).Range.Text = "";
                }

                // Подписи
                Word.Paragraph signBlock = doc.Paragraphs.Add();
                signBlock.Range.Text =
                    "\n_________________________           _________________________\n" +
                    "     Подпись владельца                     Подпись врача\n\n" +
                    "М.П.";
                signBlock.Range.Font.Size = 12;
                signBlock.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при создании паспорта: " + ex.Message);
            }


        }

    }

}





