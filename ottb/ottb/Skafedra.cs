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
    public partial class Skafedra : Form
    {
        public Skafedra()
        {
            InitializeComponent();
        }
        OleDbConnection con;    //Строка соединения с БД
        OleDbCommand SqlCom;    //Переменная для Sql запросов       
        DataTable DT;           //Таблица для хранения результатов запроса
        OleDbDataAdapter DA;    //Адаптер для заполнения таблицы после запроса      
        bool ifcon = false;     //Флаг соединения с базой данных 

        private void Skafedra_Load(object sender, EventArgs e)
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
            SqlCom = new OleDbCommand("SELECT Подразделения2.IDподразделения2, Подразделения1.НаименованиеПолное, Подразделения2.НаименованиеСокр, Подразделения2.НаименованиеПолное  " +
               " FROM Подразделения2, Подразделения1 WHERE Подразделения2.IDподразделения1=Подразделения1.IDподразделения1 ORDER BY Подразделения1.НаименованиеПолное, Подразделения2.НаименованиеСокр", con);
            SqlCom.ExecuteNonQuery();
            DA = new OleDbDataAdapter(SqlCom); //Через адаптер получаем результаты запроса
            DA.Fill(DT); // Заполняем таблицу результами
            DataGridView1.DataSource = DT;  //Привязываем DataGridView1 к источнику               
            DataGridView1.Columns[0].Visible = false; //Столбец с ID невидимый для пользователя   
            DataGridView1.Font = new Font("Times New Roman", 12);
        }
        private void ClearAll()
        {
            // Процедура очистки текстовых полей
            textBox1.Clear();
            textBox2.Clear();
            textBox6.Clear();
        }
        private bool IfNull()
        {
            //Функция обнаружения пустых полей
            if ((textBox1.Text.Trim() == "") || (textBox2.Text.Trim() == "") || (comboBox1.SelectedIndex == -1))
            {
                return true;   //Имеются пустые поля           
            }
            else
                return false;  //Пустые поля отсутствуют         
        }

        private void ShowP()
        {
            //Вывод списка  
            OleDbCommand SqlComP = new OleDbCommand("SELECT * FROM [Подразделения1]", con);
            OleDbDataReader dataReaderP = SqlComP.ExecuteReader();
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            while (dataReaderP.Read())
            {
                comboBox1.Items.Add(dataReaderP.GetValue(2));
                comboBox2.Items.Add(dataReaderP.GetValue(0));
            }
            dataReaderP.Close();
            //Исходная установка указателей списков
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
        }
        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            //Копирование строки в текстовые поля
            int i = DataGridView1.CurrentRow.Index;
            if (i >= 0)
            {
                textBox6.Text = DataGridView1[0, i].Value.ToString();   //Скрытое поле для хранения id
                comboBox1.SelectedItem = DataGridView1[1, i].Value.ToString();
                textBox1.Text = DataGridView1[2, i].Value.ToString();
                textBox2.Text = DataGridView1[3, i].Value.ToString();
            }
        }

        private void Skafedra_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ifcon) con.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Добавить
            if (!IfNull())  //Вызов функции проверки полей на пустые значения
            {
                //Пустые поля отсутствуют
                int combo1 = comboBox1.SelectedIndex;
                comboBox2.SelectedIndex = combo1;
                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "INSERT INTO [Подразделения2] (НаименованиеСокр, НаименованиеПолное, IDподразделения1) " +
                    "VALUES(@a1, @a2, @a3)";
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", textBox1.Text);
                SqlCom1.Parameters.AddWithValue("@a2", textBox2.Text);
                SqlCom1.Parameters.AddWithValue("@a3", comboBox2.SelectedItem);
                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Новое подразделение добавлено.", "ДОБАВЛЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                int combo1 = comboBox1.SelectedIndex;
                comboBox2.SelectedIndex = combo1;
                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "UPDATE Подразделения2 SET НаименованиеСокр=@a1, НаименованиеПолное=@a2, IDподразделения1=@a3 " +
                    " WHERE IDподразделения2=@a4";
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", textBox1.Text);
                SqlCom1.Parameters.AddWithValue("@a2", textBox2.Text);
                SqlCom1.Parameters.AddWithValue("@a3", comboBox2.SelectedItem);
                SqlCom1.Parameters.AddWithValue("@a4", textBox6.Text);
                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Данные о подразделении изменены.", "ДОБАВЛЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}
