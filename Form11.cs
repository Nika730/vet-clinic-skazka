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
    public partial class Form11 : Form
    {

        private BindingSource bs = new BindingSource();
        private string currentFilterMode = "";
        public Form11()
        {
            InitializeComponent();
            comboBox3.SelectedIndexChanged += comboBox3_SelectedIndexChanged;
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
                    record.id,
                    DATE(record.date_record) AS 'Дата',
                    DATE_FORMAT(record.time_record, '%H:%i') AS 'Время',
                    animals.Nickname as 'Имя животного',
                    doctors.FIO as 'Доктор',
                    services.name_services as 'Услуга'
                FROM record
                JOIN animals ON record.animal_id = animals.id
                JOIN doctors ON record.doctor_id = doctors.id
                JOIN services ON record.services_id = services.id";

                    var adapter = new MySqlDataAdapter(query, connection);
                    var table = new DataTable();
                    adapter.Fill(table);
                  
                    table.CaseSensitive = false; // для нерегистрозависимого LIKE

                    bs.DataSource = table.DefaultView;
                    dataGridView1.DataSource = bs;

                    if (dataGridView1.Columns.Contains("id"))
                        dataGridView1.Columns["id"].Visible = false;
                    // Настройка стиля заголовков
                    dataGridView1.EnableHeadersVisualStyles = false; // отключаем системный стиль
                    dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.LightBlue; // фон
                    dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.DarkBlue;  // текст
                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold); // шрифт
                    dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // выравнивание

                    if (dataGridView1.Columns["Услуга"] != null) dataGridView1.Columns["Услуга"].Width = 265;
                    if (dataGridView1.Columns["Доктор"] != null) dataGridView1.Columns["Доктор"].Width = 240;
                    if (dataGridView1.Columns["Имя животного"] != null) dataGridView1.Columns["Имя животного"].Width = 111;

               
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка подключения: " + ex.Message);
                }
            }
        }

        private void LoadRecords()
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                string query = "SELECT record.id,record.date_record as 'Дата',DATE_FORMAT(record.time_record, '%H:%i') AS 'Время',animals.Nickname as 'Имя животного',doctors.FIO as 'Доктор',services.name_services as 'Услуга' FROM record JOIN animals ON record.animal_id = animals.id JOIN doctors ON record.doctor_id = doctors.id JOIN services ON record.services_id = services.id";
                var adapter = new MySqlDataAdapter(query, conn);
                var table = new DataTable();
                adapter.Fill(table);
               
                table.CaseSensitive = false;

                bs.DataSource = table.DefaultView;
                dataGridView1.DataSource = bs;


                if (dataGridView1.Columns.Contains("id"))
                    dataGridView1.Columns["id"].Visible = false;
            }
        }

        private void Form11_Shown(object sender, EventArgs e)
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

        /////////////////////////добавление
        public class ComboItem
        {
            public int Id { get; set; }
            public string Name { get; set; }

            public ComboItem(int id, string name)
            {
                Id = id;
                Name = name;
            }

            public override string ToString()
            {
                return Name;
            }
        }
        private void LoadComboBoxes()
        {
            comboBox1.Items.Clear(); // Питомцы
            comboBox2.Items.Clear(); // Сотрудники
            comboBox3.Items.Clear(); // Услуги

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();

                // Питомцы
                MySqlCommand cmdAnimal = new MySqlCommand("SELECT id, Nickname FROM animals", conn);
                MySqlDataReader readerAnimal = cmdAnimal.ExecuteReader();
                while (readerAnimal.Read())
                {
                    comboBox1.Items.Add(new ComboItem(readerAnimal.GetInt32(0), readerAnimal.GetString(1)));
                }
                readerAnimal.Close();

                // Сотрудники
                MySqlCommand cmdDoctor = new MySqlCommand("SELECT id, FIO FROM doctors", conn);
                MySqlDataReader readerDoctor = cmdDoctor.ExecuteReader();
                while (readerDoctor.Read())
                {
                    comboBox2.Items.Add(new ComboItem(readerDoctor.GetInt32(0), readerDoctor.GetString(1)));
                }
                readerDoctor.Close();

                // Услуги
                MySqlCommand cmdService = new MySqlCommand("SELECT id, name_services FROM services", conn);
                MySqlDataReader readerService = cmdService.ExecuteReader();
                while (readerService.Read())
                {
                    comboBox3.Items.Add(new ComboItem(readerService.GetInt32(0), readerService.GetString(1)));
                }
                readerService.Close();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            // Проверка заполнения всех полей
            if (comboBox3.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, выберите услугу.");
                return;
            }

            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, выберите питомца.");
                return;
            }

            if (comboBox2.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, выберите сотрудника.");
                return;
            }

          

            // Получение значений
            string date = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string time = dateTimePicker2.Value.ToString("HH:mm:ss");

            int animalId = ((ComboItem)comboBox1.SelectedItem).Id;
            int doctorId = ((ComboItem)comboBox2.SelectedItem).Id;
            int serviceId = ((ComboItem)comboBox3.SelectedItem).Id;

            TimeSpan selectedTime = dateTimePicker2.Value.TimeOfDay;
            TimeSpan startTime = new TimeSpan(9, 0, 0);
            TimeSpan endTime = new TimeSpan(17, 40, 0);

            // Проверка диапазона
            if (selectedTime < startTime || selectedTime > endTime || selectedTime.Minutes % 15 != 0)
            {
                MessageBox.Show("Выберите время между 09:00 и 17:40 с шагом 15 минут.");
                return;
            }

            using (MySqlConnection checkConn = new MySqlConnection(connectionString))
            {
                checkConn.Open();
                string checkQuery = @"
        SELECT COUNT(*) FROM record 
        WHERE date_record = @date 
          AND time_record = @time 
          AND doctor_id = @doctor 
          AND services_id = @service";

                MySqlCommand checkCmd = new MySqlCommand(checkQuery, checkConn);
                checkCmd.Parameters.AddWithValue("@date", date);
                checkCmd.Parameters.AddWithValue("@time", time);
                checkCmd.Parameters.AddWithValue("@doctor", doctorId);
                checkCmd.Parameters.AddWithValue("@service", serviceId);

                int count = Convert.ToInt32(checkCmd.ExecuteScalar());
                if (count > 0)
                {
                    MessageBox.Show("На выбранную дату и время уже записан клиент к этому врачу на эту услугу.");
                    return;
                }
            }


            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                string query = "INSERT INTO record (date_record, time_record, animal_id, doctor_id, services_id) " +
                               "VALUES (@date, @time, @animal, @doctor, @service)";

                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@date", date);
                cmd.Parameters.AddWithValue("@time", time);
                cmd.Parameters.AddWithValue("@animal", animalId);
                cmd.Parameters.AddWithValue("@doctor", doctorId);
                cmd.Parameters.AddWithValue("@service", serviceId);

                try
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Запись успешно добавлена!");
                    LoadPetsData(); // обновить таблицу
                    ClearFields();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при добавлении: " + ex.Message);
                }
            }
        }
        private void ClearFields()
        {
            // Очистка текстовых полей
           
            textBox1.Clear(); // Вид


            // Сброс ComboBox
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;// Пол
            comboBox3.SelectedIndex = -1; // Владелец

            // Сброс даты (например, на сегодня)
            dateTimePicker1.Value = DateTime.Today;
            dateTimePicker2.Value = DateTime.Today;
        }
        ////////////////////////////////////////////СОРТИРОВКА И ФИЛЬТРАЦИЯ УСЛУГА-ВРАЧ
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Если ничего не выбрано — скрыть и выйти
            if (comboBox3.SelectedIndex == -1 || comboBox3.SelectedItem == null)
            {
                comboBox2.Visible = false;
                label4.Visible = false;
                comboBox2.Items.Clear();
                return;
            }

            // Показать контролы для врачей
            comboBox2.Visible = true;
            label4.Visible = true;

            // Очистить текущий список врачей
            comboBox2.Items.Clear();

            // Получаем выбранную услугу
            if (comboBox3.SelectedItem is ComboItem selectedService)
            {
                int selectedServiceId = selectedService.Id;

                using (MySqlConnection conn = new MySqlConnection(connectionString))
                using (MySqlCommand cmd = new MySqlCommand(@"
            SELECT d.id, d.FIO
            FROM doctors d
            JOIN doctor_services ds ON d.id = ds.doctor_id
            WHERE ds.service_id = @serviceId
            ORDER BY d.FIO;", conn))
                {
                    cmd.Parameters.AddWithValue("@serviceId", selectedServiceId);
                    conn.Open();

                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            comboBox2.Items.Add(new ComboItem(reader.GetInt32(0), reader.GetString(1)));
                        }
                    }
                }

                // Если врачей нет — спрятать и сообщить
                if (comboBox2.Items.Count == 0)
                {
                    comboBox2.Visible = false;
                    label4.Visible = false;
                    MessageBox.Show("Для выбранной услуги пока нет доступных врачей.", "Информация",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    comboBox2.SelectedIndex = 0; // выбрать первого врача по умолчанию
                }
            }
            else
            {
                // Если выбранный элемент не того типа
                comboBox2.Visible = false;
                label4.Visible = false;
            }
        }

        private DataView GetView()
        {
            var bsSrc = dataGridView1.DataSource as BindingSource;
            return bsSrc?.List as DataView;
        }


        ///////////////////УДАЛЕНИЕ
        private void button10_Click(object sender, EventArgs e)
        {

            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Пожалуйста, выберите строку для удаления.", "Удаление", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // получаем id из скрытой колонки
            int recordId = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["id"].Value);

            DialogResult confirm = MessageBox.Show("Вы уверены, что хотите удалить эту запись?",
                                                   "Подтверждение удаления",
                                                   MessageBoxButtons.YesNo,
                                                   MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                string query = "DELETE FROM record WHERE id = @id";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@id", recordId);

                try
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Запись успешно удалена.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadPetsData(); // обновляем таблицу
                    ClearFields();
                }
                catch (MySqlException ex)
                {
                    if (ex.Number == 1451)
                        MessageBox.Show("Невозможно удалить запись, так как она используется в других таблицах.", "Ошибка удаления", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                        MessageBox.Show("Ошибка при удалении: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void Form11_Load(object sender, EventArgs e)
        {
            comboBox2.Visible = false;
            label4.Visible = false;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "HH:mm";
            dateTimePicker2.ShowUpDown = true;
            LoadRecords();
            comboBox4.Items.AddRange(new string[] { "Доктор", "Имя животного", "Услуга" });
            comboBox4.SelectedIndex = 0;
        }

        private void поБуквамToolStripMenuItem_Click(object sender, EventArgs e)
        {
            currentFilterMode = "letters";
            textBox2.Visible = true;
            label6.Visible = true;
            label8.Visible = false;
            label9.Visible = false;
            comboBox4.Visible = true;
            dateTimePicker3.Visible = false;
            dateTimePicker4.Visible = false;
            dateTimePicker5.Visible = false;
            dateTimePicker6.Visible = false;
            button11.Visible = true;

        }

        private void поДатеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            currentFilterMode = "datetime";
            textBox2.Visible = false;
            comboBox4.Visible = false;
            label6.Visible = false;
            label8.Visible = true;
            label9.Visible = true;
            dateTimePicker3.Visible = true;
            dateTimePicker4.Visible = true;
            dateTimePicker5.Visible = true;
            dateTimePicker6.Visible = true;
            button11.Visible = true;

            // форматируем пикеры
          
            dateTimePicker4.Format = DateTimePickerFormat.Custom;
            dateTimePicker4.CustomFormat = "HH:mm";
            
            dateTimePicker6.Format = DateTimePickerFormat.Custom;
            dateTimePicker6.CustomFormat = "HH:mm";
        }
        private void ApplyFilter()
        {
            var view = GetView();
            if (view == null) return;

            var dt = view.Table;
            if (dt == null) return;

            dt.CaseSensitive = false;

            if (currentFilterMode == "letters")
            {
                string selectedColumn = comboBox4.SelectedItem?.ToString();
                string filterValue = textBox2.Text.Trim();

                if (!string.IsNullOrEmpty(selectedColumn) && !string.IsNullOrEmpty(filterValue))
                {
                    if (!dt.Columns.Contains(selectedColumn))
                    {
                        MessageBox.Show($"Колонка '{selectedColumn}' не найдена.");
                        return;
                    }

                    string safe = filterValue.Replace("'", "''");

                    // Приводим значение к строке, чтобы LIKE работал независимо от исходного типа
                    view.RowFilter = string.Format(
                        "CONVERT([{0}], System.String) LIKE '%{1}%'",
                        selectedColumn,
                        safe
                    );
                }
                else
                {
                    MessageBox.Show("Выберите колонку и введите значение для фильтрации!");
                }
            }
            else if (currentFilterMode == "datetime")
            {
                ApplyDateTimeFilter();
            }
        }


        private void ApplyDateTimeFilter()
        {
            var view = GetView();
            if (view == null) return;

            var dt = view.Table;
            if (dt == null) return;

            DateTime start = dateTimePicker3.Value.Date + dateTimePicker4.Value.TimeOfDay;
            DateTime end = dateTimePicker5.Value.Date + dateTimePicker6.Value.TimeOfDay;

            if (start > end)
            {
                MessageBox.Show("Начальная дата должна быть меньше или равна конечной.");
                return;
            }

            if (!dt.Columns.Contains("Дата") || !dt.Columns.Contains("Время"))
            {
                MessageBox.Show("Таблица должна содержать колонки 'Дата' и 'Время'.");
                return;
            }

            // Объединённая колонка (пересчитываем каждый раз — данные могли измениться)
            if (!dt.Columns.Contains("ДатаВремя"))
            {
                dt.Columns.Add(new DataColumn("ДатаВремя", typeof(DateTime)) { AllowDBNull = true });
            }

            foreach (DataRow row in dt.Rows)
            {
                DateTime datePart = default;
                DateTime timePart = default;

                if (dt.Columns["Дата"].DataType == typeof(DateTime))
                    datePart = ((DateTime)row["Дата"]).Date;
                else
                    DateTime.TryParse(Convert.ToString(row["Дата"]), out datePart);

                if (dt.Columns["Время"].DataType == typeof(DateTime))
                    timePart = (DateTime)row["Время"];
                else
                    DateTime.TryParse(Convert.ToString(row["Время"]), out timePart);

                if (datePart == default || timePart == default)
                    row["ДатаВремя"] = DBNull.Value;
                else
                    row["ДатаВремя"] = datePart.Date + timePart.TimeOfDay;
            }

            string startStr = start.ToString("MM/dd/yyyy HH:mm");
            string endStr = end.ToString("MM/dd/yyyy HH:mm");

            // Исключаем пустые значения
            view.RowFilter = string.Format(
                "[ДатаВремя] IS NOT NULL AND [ДатаВремя] >= #{0}# AND [ДатаВремя] <= #{1}#",
                startStr,
                endStr
            );
        }


        private void button11_Click(object sender, EventArgs e)
        {
            ApplyFilter();
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadPetsData();
            ClearFields();

            var view = GetView();
            if (view != null)
            {
                view.RowFilter = string.Empty;
                view.Sort = string.Empty;
            }

            textBox2.Clear();
            comboBox4.SelectedIndex = 0;
            textBox2.Visible = false;
            comboBox4.Visible = false;
            label6.Visible = false;
            label8.Visible = false;
            label9.Visible = false;
            dateTimePicker3.Visible = false;
            dateTimePicker4.Visible = false;
            dateTimePicker5.Visible = false;
            dateTimePicker6.Visible = false;
            button11.Visible = false;
        }

        

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string searchText = textBox1.Text.ToLower();
                UpdateMatchSort(searchText);
                e.Handled = true;
            }
        }

        private void UpdateMatchSort(string searchText)
        {
            var view = bs.DataSource as DataView;
            if (view == null) return;

            var dt = view.Table;

            // Добавляем колонку MatchSort, если её нет
            if (!dt.Columns.Contains("MatchSort"))
            {
                dt.Columns.Add(new DataColumn("MatchSort", typeof(int)));
            }

            // Получаем id выделенной строки
            int? currentId = null;
            if (dataGridView1.CurrentRow?.Cells["id"]?.Value != null)
            {
                currentId = Convert.ToInt32(dataGridView1.CurrentRow.Cells["id"].Value);
            }

            foreach (DataRow row in dt.Rows)
            {
                string combined = string.Join(" ",
                    row["Дата"]?.ToString()?.ToLower() ?? "",
                    row["Время"]?.ToString()?.ToLower() ?? "",
                    row["Имя животного"]?.ToString()?.ToLower() ?? "",
                    row["Доктор"]?.ToString()?.ToLower() ?? "",
                    row["Услуга"]?.ToString()?.ToLower() ?? "");

                bool isMatch = !string.IsNullOrEmpty(searchText) && combined.Contains(searchText);

                int rowId = dt.Columns.Contains("id") && row["id"] != DBNull.Value
                    ? Convert.ToInt32(row["id"])
                    : -1;

                // Логика сортировки:
                // 2 — совпадает
                // -1 — выделенная строка, не совпадает
                // 0 — остальные
                if (isMatch)
                {
                    row["MatchSort"] = 2;
                }
                else if (currentId.HasValue && rowId == currentId.Value)
                {
                    row["MatchSort"] = -1;
                }
                else
                {
                    row["MatchSort"] = 0;
                }
            }

            // Применяем сортировку
            view.Sort = "MatchSort DESC";
            bs.ResetBindings(false);
            if (dataGridView1.Columns.Contains("MatchSort"))
                dataGridView1.Columns["MatchSort"].Visible = false;
        }








        private void textBox1_TextChanged(object sender, EventArgs e)
        
        {
            string searchText = textBox1.Text.ToLower();

            // Сброс/установка подсветки
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (string.IsNullOrEmpty(searchText))
                    {
                        cell.Style.BackColor = Color.White;
                        continue;
                    }

                    if (cell.Value != null && cell.Value.ToString().ToLower().Contains(searchText))
                        cell.Style.BackColor = Color.LightBlue;
                    else
                        cell.Style.BackColor = Color.White;
                }
            }

            // Обновляем колонку сортировки, но выделенная строка пока не исключается
            EnsureMatchSortColumn();
            UpdateMatchSort(searchText, excludeCurrentRowFromSort: false);
        }

        
        private void EnsureMatchSortColumn()
        {
            var view = bs.DataSource as DataView;
            if (view == null) return;

            var dt = view.Table;
            if (!dt.Columns.Contains("MatchSort"))
            {
                dt.Columns.Add(new DataColumn("MatchSort", typeof(int)));
                foreach (DataRow r in dt.Rows) r["MatchSort"] = 0;
            }
        }

        private void UpdateMatchSort(string searchText, bool excludeCurrentRowFromSort)
        {
            var dt = bs.DataSource as DataTable;
            if (dt == null) return;

            // Определяем текущую выделенную строку (по id)
            int? currentId = null;
            if (excludeCurrentRowFromSort && dataGridView1.CurrentRow != null)
            {
                var cell = dataGridView1.CurrentRow.Cells["id"];
                if (cell != null && cell.Value != null)
                    currentId = Convert.ToInt32(cell.Value);
            }

            foreach (DataRow row in dt.Rows)
            {
                // Пропускаем выделенную строку, если нужно
                if (excludeCurrentRowFromSort && currentId.HasValue && dt.Columns.Contains("id"))
                {
                    int rid = Convert.ToInt32(row["id"]);
                    if (rid == currentId.Value)
                    {
                        row["MatchSort"] = 0;
                        continue;
                    }
                }

                string date = row["Дата"]?.ToString()?.ToLower() ?? "";
                string time = row["Время"]?.ToString()?.ToLower() ?? "";
                string pet = row["Имя животного"]?.ToString()?.ToLower() ?? "";
                string doc = row["Доктор"]?.ToString()?.ToLower() ?? "";
                string srv = row["Услуга"]?.ToString()?.ToLower() ?? "";

                bool match = !string.IsNullOrEmpty(searchText) &&
                             (date.Contains(searchText) || time.Contains(searchText) ||
                              pet.Contains(searchText) || doc.Contains(searchText) || srv.Contains(searchText));

                row["MatchSort"] = match ? 1 : 0;
            }

            // Сортируем: совпадения сверху, остальное ниже
            dt.DefaultView.Sort = "MatchSort DESC";
        }

        private void wordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите запись в таблице!");
                return;
            }

            var row = dataGridView1.SelectedRows[0];

            // Обработка даты
            string rawDate = row.Cells["Дата"].Value?.ToString();
            DateTime parsedDate;
            string date = DateTime.TryParse(rawDate, out parsedDate)
                ? parsedDate.ToString("dd.MM.yyyy")
                : rawDate;

            string time = row.Cells["Время"].Value?.ToString();
            string animal = row.Cells["Имя животного"].Value?.ToString();
            string doctor = row.Cells["Доктор"].Value?.ToString();
            string service = row.Cells["Услуга"].Value?.ToString();

            try
            {
                Word.Application wordApp = new Word.Application();
                wordApp.Visible = true;
                Word.Document doc = wordApp.Documents.Add();

                // Настройка размера страницы (компактный формат)
                doc.PageSetup.PageWidth = wordApp.CentimetersToPoints(10);  // ширина ~10 см
                doc.PageSetup.PageHeight = wordApp.CentimetersToPoints(15); // высота ~15 см

                // Поля
                doc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1);
                doc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1);
                doc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(1);
                doc.PageSetup.RightMargin = wordApp.CentimetersToPoints(1);

                // Заголовок
                Word.Paragraph header = doc.Paragraphs.Add();
                header.Range.Text = "ТАЛОН НА ПРИЁМ";
                header.Range.Font.Size = 16;
                header.Range.Font.Bold = 1;
                header.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                header.Range.InsertParagraphAfter();

                // Дата
                Word.Paragraph pDate = doc.Paragraphs.Add();
                pDate.Range.Text = "Дата: " + date;
                pDate.Range.Font.Size = 12;
                pDate.Range.InsertParagraphAfter();

                // Время
                Word.Paragraph pTime = doc.Paragraphs.Add();
                pTime.Range.Text = "Время: " + time;
                pTime.Range.Font.Size = 12;
                pTime.Range.InsertParagraphAfter();

                // Пациент
                Word.Paragraph pAnimal = doc.Paragraphs.Add();
                pAnimal.Range.Text = "Пациент: " + animal;
                pAnimal.Range.Font.Size = 12;
                pAnimal.Range.InsertParagraphAfter();

                // Доктор
                Word.Paragraph pDoctor = doc.Paragraphs.Add();
                pDoctor.Range.Text = "Доктор: " + doctor;
                pDoctor.Range.Font.Size = 12;
                pDoctor.Range.InsertParagraphAfter();

                // Услуга
                Word.Paragraph pService = doc.Paragraphs.Add();
                pService.Range.Text = "Услуга: " + service;
                pService.Range.Font.Size = 12;
                pService.Range.InsertParagraphAfter();

                // Подпись
                Word.Paragraph footer = doc.Paragraphs.Add();
                footer.Range.Text = "\nПожалуйста, приходите вовремя!";
                footer.Range.Font.Italic = 1;
                footer.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                footer.Range.InsertParagraphAfter();

                // Вставка фото под текст
                string photoPath = @"D:\logo.png"; // путь к фото
                if (System.IO.File.Exists(photoPath))
                {
                    Word.Paragraph imgParagraph = doc.Paragraphs.Add();
                    imgParagraph.Range.InsertParagraphAfter();

                    Word.InlineShape photo = imgParagraph.Range.InlineShapes.AddPicture(photoPath);
                    photo.Width = 170;
                    photo.Height = 120;
                    imgParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при формировании талона: " + ex.Message);
            }



        }

        private void договорToolStripMenuItem_Click(object sender, EventArgs e)
        {
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
                title.Range.Text = "ДОГОВОР НА ВЕТЕРИНАРНОЕ ОБСЛУЖИВАНИЕ";
                title.Range.Font.Size = 14;
                title.Range.Font.Bold = 1;
                title.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                title.Range.InsertParagraphAfter();

                // Вводная часть
                Word.Paragraph intro = doc.Paragraphs.Add();
                intro.Range.Text =
                    "г. Гомель\n" +
                    $"Дата: {DateTime.Now:dd.MM.yyyy}\n\n" +
                    "Настоящий договор заключён между ветеринарной клиникой «Сказка» (далее — Исполнитель) и владельцем животного (далее — Заказчик).\n";
                intro.Range.Font.Size = 12;
                intro.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                intro.Range.InsertParagraphAfter();

                // Основные условия
                Word.Paragraph terms = doc.Paragraphs.Add();
                terms.Range.Text =
                    "1. Предмет договора\n" +
                    "1.1. Исполнитель обязуется предоставить ветеринарные услуги, а Заказчик — оплатить их.\n" +
                    "2. Обязанности сторон\n" +
                    "2.1. Исполнитель предоставляет услуги квалифицированно и своевременно.\n" +
                    "2.2. Заказчик предоставляет достоверную информацию о состоянии животного.\n" +
                    "3. Стоимость и порядок оплаты\n" +
                    "3.1. Стоимость услуг определяется согласно прайс-листу клиники.\n" +
                    "3.2. Оплата производится наличными или по безналичному расчёту.\n" +
                    "4. Ответственность сторон\n" +
                    "4.1. Стороны несут ответственность согласно действующему законодательству Республики Беларусь.\n" +
                    "5. Прочие условия\n" +
                    "5.1. Договор вступает в силу с момента подписания и действует до полного исполнения обязательств.\n";
                terms.Range.Font.Size = 12;
                terms.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                terms.Range.InsertParagraphAfter();

                // Подписи
                // Вставка таблицы для подписей
                Word.Table signatureTable = doc.Tables.Add(doc.Paragraphs.Add().Range, 3, 2);
                signatureTable.Range.Font.Size = 12;
                signatureTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                signatureTable.Columns[1].Width = wordApp.CentimetersToPoints(8f);
                signatureTable.Columns[2].Width = wordApp.CentimetersToPoints(8f);

                // Строка 1 — линии
                signatureTable.Cell(1, 1).Range.Text = "__________________________";
                signatureTable.Cell(1, 2).Range.Text = "__________________________";

                // Строка 2 — ФИО
                signatureTable.Cell(2, 1).Range.Text = "Заказчик (ФИО)";
                signatureTable.Cell(2, 2).Range.Text = "Исполнитель (ФИО)";

                // Строка 3 — подпись
                signatureTable.Cell(3, 1).Range.Text = "Подпись:";
                signatureTable.Cell(3, 2).Range.Text = "Подпись:";

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при создании договора: " + ex.Message);
            }



        }
    }

}









