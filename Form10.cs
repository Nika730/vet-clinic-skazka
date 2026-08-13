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

namespace Pets
{
    public partial class Form10 : Form
    {
        private static string connectionString = "server=localhost;user=root;password=admin;database=pet;";

        public Form10()
        {
            InitializeComponent();
        }
        ////////////ПЕРЕЙТИ НА ВХОД
        private void button2_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Hide();
            this.Close();
        }
        ///////////ЗАРЕГИСТРИРОВАТЬСЯ
        private void button1_Click(object sender, EventArgs e)
        {
            // допустим, у тебя есть TextBox'ы: textBoxNick и textBoxPass
            string nick = textBox1.Text.Trim();
            string pass = textBox2.Text.Trim();

            if (string.IsNullOrEmpty(nick) || string.IsNullOrEmpty(pass))
            {
                MessageBox.Show("Введите логин и пароль!");
                return;
            }
            // Проверка длины
            if (nick.Length > 25)
            {
                MessageBox.Show("Логин не должен превышать 25 символов!");
                return;
            }
            if (pass.Length < 4 || pass.Length > 9)
            {
                MessageBox.Show("Пароль должен содержать от 4 до 9 символов!");
                return;
            }

            // Проверка допустимых символов (только латиница и цифры)
            Regex regex = new Regex("^[a-zA-Z0-9]+$");
            if (!regex.IsMatch(nick))
            {
                MessageBox.Show("Логин может содержать только английские буквы и цифры!");
                return;
            }
            if (!regex.IsMatch(pass))
            {
                MessageBox.Show("Пароль может содержать только английские буквы и цифры!");
                return;
            }
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    string checkQuery = "SELECT COUNT(*) FROM users WHERE nick = @nick";
                    using (MySqlCommand checkCmd = new MySqlCommand(checkQuery, conn))
                    {
                        checkCmd.Parameters.AddWithValue("@nick", nick);
                        int count = Convert.ToInt32(checkCmd.ExecuteScalar());
                        if (count > 0)
                        {
                            MessageBox.Show("Такой логин уже существует!");
                            return;
                        }
                    }
                    // Добавление нового пользователя
                    string query = "INSERT INTO users (nick, levl, pass) VALUES (@nick, @levl, @pass)";
                    using (MySqlCommand cmd = new MySqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@nick", nick);
                        cmd.Parameters.AddWithValue("@levl", "Пользователь");
                        cmd.Parameters.AddWithValue("@pass", pass);

                        cmd.ExecuteNonQuery();
                    }

                    // Очистка полей после успешной регистрации
                    textBox1.Clear();
                    textBox2.Clear();

                    MessageBox.Show("Регистрация прошла успешно!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка: " + ex.Message);
                }
            }
        }

        private void Form10_Load(object sender, EventArgs e)
        {
            textBox1.MaxLength = 25; // максимум 25 символов
            textBox2.MaxLength = 9;  // максимум 9 символов
        }
         
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) &&
        !(e.KeyChar >= 'a' && e.KeyChar <= 'z') &&
        !(e.KeyChar >= 'A' && e.KeyChar <= 'Z') &&
        !(e.KeyChar >= '0' && e.KeyChar <= '9'))
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) &&
         !(e.KeyChar >= 'a' && e.KeyChar <= 'z') &&
         !(e.KeyChar >= 'A' && e.KeyChar <= 'Z') &&
         !(e.KeyChar >= '0' && e.KeyChar <= '9'))
            {
                e.Handled = true;
            }
        }
    }
}

