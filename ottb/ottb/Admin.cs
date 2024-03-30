using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//-------------- Подключить библиотеку для работы с БД ---------------------------------
using System.Data.OleDb;

namespace ottb
{
    public partial class Admin : Form
    {
        public Admin()
        {
            InitializeComponent();
        }
        OleDbConnection con;    //Строка соединения с БД
        OleDbCommand SqlCom;    //Переменная для Sql запросов       
        DataTable DT;           //Таблица для хранения результатов запроса
        OleDbDataAdapter DA;    //Адаптер для заполнения таблицы после запроса      
        bool ifcon = false;     //Флаг соединения с базой данных    

        private void ShowP()
        {
            //Вывод списка  
             
            OleDbCommand SqlComP = new OleDbCommand("SELECT * FROM [Подразделения2]", con);
            OleDbDataReader dataReaderP = SqlComP.ExecuteReader();
            comboBox3.Items.Clear();
            comboBox4.Items.Clear();
            while (dataReaderP.Read())
            {
                comboBox3.Items.Add(dataReaderP.GetValue(2));
                comboBox4.Items.Add(dataReaderP.GetValue(0));
            }
            dataReaderP.Close();
            //Исходная установка указателей списков
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;

        }
        private void Admin_Load(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=ottb.accdb");
                con.Open();     //Открыть базу данных
                ifcon = true;   //Флаг поднят. Соединение с базой данных прошло успешно.  
                ShowP();
                ShowList();
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "ОШИБКА ДОСТУПА К БАЗЕ ДАННЫХ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowList()
        {
            //Процедура вывода списка в таблицу DataGridView1           
            DT = new DataTable();  //Создаем заново таблицу               
            // Указываем строку запроса и привязываем к соединению
            SqlCom = new OleDbCommand("SELECT * FROM [Пользователи] ORDER BY [Роль]", con);
            SqlCom.ExecuteNonQuery();
            DA = new OleDbDataAdapter(SqlCom); //Через адаптер получаем результаты запроса
            DA.Fill(DT); // Заполняем таблицу результами
            DataGridView1.DataSource = DT;  //Привязываем DataGridView1 к источнику данных             
            DataGridView1.Columns[0].Visible = false; //Столбец с ID невидимый для пользователя   
            DataGridView1.Font = new Font("Times New Roman", 12);
        }
        private void ClearAll()
        {
            // Процедура очистки текстовых полей
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = 1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = 1;

        }
        private bool IfNull()
        {
            //Функция обнаружения пустых полей
            if ((textBox2.Text.Trim() == "") || (textBox3.Text.Trim() == "") || (textBox4.Text.Trim() == "")||
                (comboBox1.SelectedIndex == -1) || (comboBox2.SelectedIndex == -1) || (comboBox3.SelectedIndex == -1))
            {
                return true;   //Имеются пустые поля           
            }
            else
                return false;  //Пустые поля отсутствуют         
        }

        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            //Копирование строки в текстовые поля
            int i = DataGridView1.CurrentRow.Index;
            if (i >= 0)
            {
                textBox6.Text = DataGridView1[0, i].Value.ToString();   //Скрытое поле для хранения id
                comboBox3.SelectedItem = DataGridView1[1, i].Value.ToString();
                textBox2.Text = DataGridView1[2, i].Value.ToString();    
                textBox3.Text = DataGridView1[3, i].Value.ToString();
                textBox4.Text = DataGridView1[4, i].Value.ToString();
                comboBox1.SelectedItem = DataGridView1[5, i].Value.ToString();
                comboBox2.SelectedItem = DataGridView1[6, i].Value.ToString();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            String s = "_QWERTYUIOPASDFGHJKLZXCVBNM%123456789qwertyuiopasdfghjklzxcvbnm%";
            String  psw = "";
            int i;
            Random  randObj = new Random();
            for (int j = 1; j <= 8; j++)
            {
                i = randObj.Next(0, 60);
                psw = psw + s[i];
            }
            textBox4.Text = psw;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            String  fio;
            String  pasw;
            fio = Convert.ToString(textBox1.Text);
            pasw = Convert.ToString(textBox4.Text);

            textBox5.Text = fio + "\r\n" + "Ваш временный пароль: " + pasw+"\r\n";
            string MyFile1 = "Пароли.txt";
            var MyWrite = new System.IO.StreamWriter(MyFile1, true);
            MyWrite.WriteLine(textBox5.Text, true);
            MyWrite.Close();             
            try
            {
                System.Diagnostics.Process.Start("Notepad", "Пароли.txt");
            }
            catch	
            {
                MessageBox.Show("Файл с паролями не найден!", "ОШИБКА ЧТЕНИЯ ФАЙЛА", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Admin_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ifcon) con.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Добавить
            if (!IfNull())  //Вызов функции проверки полей на пустые значения
            {
                //Пустые поля отсутствуют
                int combo = comboBox3.SelectedIndex;
                comboBox4.SelectedIndex = combo;
                textBox7.Text = Convert.ToString(comboBox4.SelectedItem);
                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "INSERT INTO [Пользователи] (Подразделение, Руководитель, Логин, Пароль, Роль, Блокирован, IDподразделения2) " +
                    "VALUES(@a1, @a2, @a3, @a4, @a5, @a6, @a7)";
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", comboBox3.SelectedItem);  
                SqlCom1.Parameters.AddWithValue("@a2", textBox2.Text);   
                SqlCom1.Parameters.AddWithValue("@a3", textBox3.Text);   
                SqlCom1.Parameters.AddWithValue("@a4", textBox4.Text);
                SqlCom1.Parameters.AddWithValue("@a5", comboBox1.SelectedItem);   
                SqlCom1.Parameters.AddWithValue("@a6", comboBox2.SelectedItem);
                SqlCom1.Parameters.AddWithValue("@a7", textBox7.Text);
                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Новый пользователь добавлен.", "ДОБАВЛЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("ПУСТЫЕ ПОЛЯ НЕ ДОПУСТИМЫ!", "КОНТРОЛЬ ДАННЫХ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Изменить
            if (!IfNull())  //Вызов функции проверки полей на пустые значения
            {
                //Пустые поля отсутствуют
                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "UPDATE [Пользователи] SET Подразделение=@a1, Руководитель=@a2, Логин=@a3, Пароль=@a4, Роль=@a5, Блокирован=@a6, IDподразделения2=@a7 WHERE IDпользователя=@a8"; 
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", comboBox3.SelectedItem);
                SqlCom1.Parameters.AddWithValue("@a2", textBox2.Text);
                SqlCom1.Parameters.AddWithValue("@a3", textBox3.Text);
                SqlCom1.Parameters.AddWithValue("@a4", textBox4.Text);
                SqlCom1.Parameters.AddWithValue("@a5", comboBox1.SelectedItem);
                SqlCom1.Parameters.AddWithValue("@a6", comboBox2.SelectedItem);
                SqlCom1.Parameters.AddWithValue("@a7", textBox7.Text);
                SqlCom1.Parameters.AddWithValue("@a8", textBox6.Text);
                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Изменения сохранены.", "ИЗМЕНЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("ПУСТЫЕ ПОЛЯ НЕ ДОПУСТИМЫ!", "КОНТРОЛЬ ДАННЫХ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Удалить
            if (textBox6.Text.Trim() == "")
            {
                MessageBox.Show("Выберите строку для удаления!",
                "ОШИБКА В ОПЕРАЦИИ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "DELETE FROM [Пользователи] WHERE IDпользователя = @id";
                SqlCom1.Parameters.Clear();
                SqlCom1.Parameters.AddWithValue("@id", textBox6.Text);
                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Запись удалена.", "УДАЛЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ClearAll();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Spravka f = new Spravka();
            f.Show();
        }
    }
}
