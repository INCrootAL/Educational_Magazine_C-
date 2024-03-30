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
    public partial class Sinstitut : Form
    {
        public Sinstitut()
        {
            InitializeComponent();
        }
        OleDbConnection con;    //Строка соединения с БД
        OleDbCommand SqlCom;    //Переменная для Sql запросов       
        DataTable DT;           //Таблица для хранения результатов запроса
        OleDbDataAdapter DA;    //Адаптер для заполнения таблицы после запроса      
        bool ifcon = false;     //Флаг соединения с базой данных   

        private void Sinstitut_Load(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=ottb.accdb");
                con.Open();     //Открыть базу данных
                ifcon = true;   //Флаг поднят. Соединение с базой данных прошло успешно. 
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
            SqlCom = new OleDbCommand("SELECT * FROM [Подразделения1] ORDER BY [НаименованиеСокр]", con);
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
            textBox6.Clear();      
        }
        private bool IfNull()
        {
            //Функция обнаружения пустых полей
            if ((textBox1.Text.Trim() == "") || (textBox2.Text.Trim() == ""))
            {
                return true;   //Имеются пустые поля           
            }
            else
                return false;  //Пустые поля отсутствуют         
        }

        private void Sinstitut_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ifcon) con.Close();
        }

        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            //Копирование строки в текстовые поля
            int i = DataGridView1.CurrentRow.Index;
            if (i >= 0)
            {
                textBox6.Text = DataGridView1[0, i].Value.ToString();   //Скрытое поле для хранения id
                textBox1.Text = DataGridView1[1, i].Value.ToString();
                textBox2.Text = DataGridView1[2, i].Value.ToString();                 
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Добавить
            if (!IfNull())  //Вызов функции проверки полей на пустые значения
            {
                //Пустые поля отсутствуют
                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "INSERT INTO [Подразделения1] (НаименованиеСокр, НаименованиеПолное) " +
                    "VALUES(@a1, @a2)";
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", textBox1.Text);
                SqlCom1.Parameters.AddWithValue("@a2", textBox2.Text);                 
                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Новая запись добавлена.", "ДОБАВЛЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                SqlCom1.CommandText = "UPDATE [Подразделения1] SET НаименованиеСокр=@a1, НаименованиеПолное=@a2 WHERE IDподразделения1=@a7";
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", textBox1.Text);
                SqlCom1.Parameters.AddWithValue("@a2", textBox2.Text);                
                SqlCom1.Parameters.AddWithValue("@a7", textBox6.Text);
                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Изменения сохранены.", "ИЗМЕНЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("ПУСТЫЕ ПОЛЯ НЕ ДОПУСТИМЫ!", "КОНТРОЛЬ ДАННЫХ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ClearAll();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}
