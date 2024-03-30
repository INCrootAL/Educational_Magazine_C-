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
    public partial class Shtat : Form
    {
        public Shtat()
        {
            InitializeComponent();
        }
        OleDbConnection con;    //Строка соединения с БД
        OleDbCommand SqlCom;    //Переменная для Sql запросов       
        DataTable DT;           //Таблица для хранения результатов запроса
        OleDbDataAdapter DA;    //Адаптер для заполнения таблицы после запроса      
        bool ifcon = false;     //Флаг соединения с базой данных 
        int npid = 0;
        private void Shtat_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ifcon) con.Close();
        }

        private void ClearAll()
        {
            // Процедура очистки текстовых полей
            textBox1.Clear();
            comboBox1.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
        }
        private void ShowList()
        {
            //Процедура вывода списка в таблицу DataGridView1           
            DT = new DataTable();  //Создаем заново таблицу               
            // Указываем строку запроса и привязываем к соединению
            SqlCom = new OleDbCommand("SELECT Штат.IDработника, Штат.ФИО, Штат.ДатаРождения, Должности.Наименование, Штат.Блокирован " +
                " FROM Штат, Должности WHERE Штат.IDдолжности=Должности.IDдолжности AND IDподразделения2=" + npid + " ORDER BY ФИО", con);

            SqlCom.ExecuteNonQuery();
            DA = new OleDbDataAdapter(SqlCom); //Через адаптер получаем результаты запроса
            DA.Fill(DT); // Заполняем таблицу результами
            DataGridView1.DataSource = DT;  //Привязываем DataGridView1 к источнику               
            DataGridView1.Columns[0].Visible = false; //Столбец с ID невидимый для пользователя             
            DataGridView1.Font = new Font("Times New Roman", 12);
        }

        private void ShowP()
        {
            SqlCom = new OleDbCommand("SELECT * FROM [Временная1]", con);
            OleDbDataReader dataReaderV = SqlCom.ExecuteReader();
            String pid = "";
            while (dataReaderV.Read())  //Пока не конец виртуальной таблицы
            {
                pid = Convert.ToString(dataReaderV.GetValue(0));
            }
            dataReaderV.Close();    //Закрыть объект чтения
            npid = Convert.ToInt32(pid);

            OleDbCommand SqlComP = new OleDbCommand("SELECT * FROM [Должности]", con);
            OleDbDataReader dataReaderP = SqlComP.ExecuteReader();
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            while (dataReaderP.Read())
            {
                comboBox1.Items.Add(dataReaderP.GetValue(1));
                comboBox2.Items.Add(dataReaderP.GetValue(0));
            }
            dataReaderP.Close();
        }

        private bool IfNull()
        {
            //Функция обнаружения пустых полей
            if ((textBox1.Text.Trim() == "") || (comboBox1.SelectedIndex == -1) || (comboBox3.SelectedIndex == -1))
            {
                return true;   //Имеются пустые поля           
            }
            else
                return false;  //Пустые поля отсутствуют         
        }
        private void Shtat_Load(object sender, EventArgs e)
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

        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            //Копирование строки в текстовые поля
            if (DataGridView1.RowCount > 1)
            {
                int i = DataGridView1.CurrentRow.Index;
                if (i >= 0)
                {
                    textBox6.Text = DataGridView1[0, i].Value.ToString();   //Скрытое поле для хранения id
                    textBox1.Text = DataGridView1[1, i].Value.ToString();
                    dateTimePicker1.Value = Convert.ToDateTime(DataGridView1[2, i].Value);
                    comboBox1.SelectedItem = DataGridView1[3, i].Value.ToString();
                    comboBox3.SelectedItem = DataGridView1[4, i].Value.ToString();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Добавить
            if (!IfNull())  //Вызов функции проверки полей на пустые значения
            {
                //Пустые поля отсутствуют
                DateTime dat = dateTimePicker1.Value;
                //Приведение даты к формату dd.mm.yyyy

                int d = dat.Day;
                int m = dat.Month;
                int y = dat.Year;
                String d1 = Convert.ToString(d);
                String m1 = Convert.ToString(m);
                if (d < 10)
                    d1 = "0" + d1;
                if (m < 10)
                    m1 = "0" + m1;
                String dats = d1 + "." + m1 + "." + Convert.ToString(y);
                int combo1 = comboBox1.SelectedIndex;
                comboBox2.SelectedIndex = combo1;
                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "INSERT INTO [Штат] (IDподразделения2, ФИО, ДатаРождения, IDдолжности, Блокирован) " +
                    "VALUES(@a1, @a2, @a3, @a4, @a5)";
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", npid);
                SqlCom1.Parameters.AddWithValue("@a2", textBox1.Text);
                SqlCom1.Parameters.AddWithValue("@a3", dats);
                SqlCom1.Parameters.AddWithValue("@a4", comboBox2.SelectedItem);
                SqlCom1.Parameters.AddWithValue("@a5", comboBox3.SelectedItem);
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
                DateTime dat = dateTimePicker1.Value;
                //Приведение даты к формату dd.mm.yyyy

                int d = dat.Day;
                int m = dat.Month;
                int y = dat.Year;
                String d1 = Convert.ToString(d);
                String m1 = Convert.ToString(m);
                if (d < 10)
                    d1 = "0" + d1;
                if (m < 10)
                    m1 = "0" + m1;
                String dats = d1 + "." + m1 + "." + Convert.ToString(y);
                int combo1 = comboBox1.SelectedIndex;
                comboBox2.SelectedIndex = combo1;
                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "UPDATE [Штат] SET IDподразделения2=@a1, ФИО=@a2, ДатаРождения=@a3, IDдолжности=@a4, Блокирован=@a5 WHERE IDработника =@a6";
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", npid);
                SqlCom1.Parameters.AddWithValue("@a2", textBox1.Text);
                SqlCom1.Parameters.AddWithValue("@a3", dats);
                SqlCom1.Parameters.AddWithValue("@a4", comboBox2.SelectedItem);
                SqlCom1.Parameters.AddWithValue("@a5", comboBox3.SelectedItem);
                SqlCom1.Parameters.AddWithValue("@a6", textBox6.Text);
                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Запись изменена.", "ДОБАВЛЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("ПУСТЫЕ ПОЛЯ НЕ ДОПУСТИМЫ!", "КОНТРОЛЬ ДАННЫХ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ClearAll();
        }
    }
}
