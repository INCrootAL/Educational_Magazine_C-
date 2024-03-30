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
    public partial class Jurnal2 : Form
    {
        public Jurnal2()
        {
            InitializeComponent();
        }
        OleDbConnection con;    //Строка соединения с БД
        OleDbCommand SqlCom;    //Переменная для Sql запросов       
        DataTable DT;           //Таблица для хранения результатов запроса
        OleDbDataAdapter DA;    //Адаптер для заполнения таблицы после запроса      
        bool ifcon = false;     //Флаг соединения с базой данных 

        private void ShowList()
        {
            //Процедура вывода списка в таблицу DataGridView1           
            DT = new DataTable();  //Создаем заново таблицу               
            // Указываем строку запроса и привязываем к соединению
            if (radioButton1.Checked)
                SqlCom = new OleDbCommand("SELECT * FROM Журнал2 ORDER BY Дата DESC", con);
            if (radioButton2.Checked)
                SqlCom = new OleDbCommand("SELECT * FROM Журнал2 ORDER BY Работник", con);
            if (radioButton3.Checked)
                SqlCom = new OleDbCommand("SELECT * FROM Журнал2 ORDER BY Подразделение2", con);

            SqlCom.ExecuteNonQuery();
            DA = new OleDbDataAdapter(SqlCom); //Через адаптер получаем результаты запроса
            DA.Fill(DT); // Заполняем таблицу результами
            DataGridView1.DataSource = DT;  //Привязываем DataGridView1 к источнику               
            DataGridView1.Columns[0].Visible = false; //Столбец с ID невидимый для пользователя   
            DataGridView1.Columns[16].Visible = false;
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
            textBox7.Clear();
            textBox8.Clear();
            textBox9.Clear();
            comboBox1.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            //comboBox5.SelectedIndex = -1;

        }
        private bool IfNull()
        {
            //Функция обнаружения пустых полей
            if ((textBox1.Text.Trim() == "") || (textBox2.Text.Trim() == "") || (textBox3.Text.Trim() == "") || (textBox4.Text.Trim() == "")||
                (textBox5.Text.Trim() == "") || (textBox7.Text.Trim() == "") || (textBox8.Text.Trim() == "") || (textBox9.Text.Trim() == "")||
                (textBox10.Text.Trim() == "") || (comboBox1.SelectedIndex == -1) || (comboBox3.SelectedIndex == -1))
            {
                return true;   //Имеются пустые поля           
            }
            else
                return false;  //Пустые поля отсутствуют         
        }

        private void ShowP()
        {
            //Вывод списка  
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
            //Исходная установка указателей списков
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            SqlComP = new OleDbCommand("SELECT * FROM [Подразделения2]", con);
            dataReaderP = SqlComP.ExecuteReader();
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

        private void Jurnal2_Load(object sender, EventArgs e)
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

        private void Jurnal2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ifcon) con.Close();
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
                    textBox1.Text = DataGridView1[2, i].Value.ToString();
                    comboBox1.SelectedItem = DataGridView1[3, i].Value.ToString();
                    comboBox3.SelectedItem = DataGridView1[4, i].Value.ToString();
                    textBox11.Text = DataGridView1[5, i].Value.ToString();
                    dateTimePicker1.Value = Convert.ToDateTime(DataGridView1[6, i].Value);
                    textBox3.Text = DataGridView1[7, i].Value.ToString();
                    textBox2.Text = DataGridView1[8, i].Value.ToString();
                    textBox4.Text = DataGridView1[9, i].Value.ToString();
                    dateTimePicker2.Value = Convert.ToDateTime(DataGridView1[10, i].Value);
                    textBox5.Text = DataGridView1[11, i].Value.ToString();
                    textBox7.Text = DataGridView1[12, i].Value.ToString();
                    textBox8.Text = DataGridView1[13, i].Value.ToString();
                    textBox9.Text = DataGridView1[14, i].Value.ToString();
                    textBox10.Text = DataGridView1[15, i].Value.ToString();
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

                DateTime dat2 = dateTimePicker2.Value;
                //Приведение даты к формату dd.mm.yyyy

                int d2 = dat2.Day;
                int m2 = dat2.Month;
                int y2 = dat2.Year;
                String d21 = Convert.ToString(d2);
                String m21 = Convert.ToString(m2);
                if (d2 < 10)
                    d21 = "0" + d21;
                if (m2 < 10)
                    m21 = "0" + m21;
                String dats2 = d21 + "." + m21 + "." + Convert.ToString(y2);               

                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "INSERT INTO Журнал2 (IDжурнала, Работник, Должность, Подразделение2, Место, Дата, Время, Обстоятельства, " +
                        " Причины, МедпунктДата, МедпунктВремя, Диагноз, ПринятыеМеры, Последствия, Инструктор) "+
                        " VALUES(@a1, @a2, @a3, @a4, @a5, @a6, @a7, @a8, @a9, @a10, @a11, @a12, @a13, @a14, @a15)";
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", 2);
                SqlCom1.Parameters.AddWithValue("@a2", textBox1.Text);
                SqlCom1.Parameters.AddWithValue("@a3", Convert.ToString(comboBox1.SelectedItem));
                SqlCom1.Parameters.AddWithValue("@a4", Convert.ToString(comboBox3.SelectedItem));
                SqlCom1.Parameters.AddWithValue("@a5", textBox11.Text);
                SqlCom1.Parameters.AddWithValue("@a6", dats);
                SqlCom1.Parameters.AddWithValue("@a7", textBox3.Text);
                SqlCom1.Parameters.AddWithValue("@a8", textBox2.Text);
                SqlCom1.Parameters.AddWithValue("@a9", textBox4.Text);
                SqlCom1.Parameters.AddWithValue("@a10", dats2);
                SqlCom1.Parameters.AddWithValue("@a11", textBox5.Text);
                SqlCom1.Parameters.AddWithValue("@a12", textBox7.Text);
                SqlCom1.Parameters.AddWithValue("@a13", textBox8.Text);
                SqlCom1.Parameters.AddWithValue("@a14", textBox9.Text);
                SqlCom1.Parameters.AddWithValue("@a15", textBox10.Text);


                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Запись добавлена.", "ДОБАВЛЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                DateTime dat2 = dateTimePicker2.Value;
                //Приведение даты к формату dd.mm.yyyy

                int d2 = dat2.Day;
                int m2 = dat2.Month;
                int y2 = dat2.Year;
                String d21 = Convert.ToString(d2);
                String m21 = Convert.ToString(m2);
                if (d2 < 10)
                    d21 = "0" + d21;
                if (m2 < 10)
                    m21 = "0" + m21;
                String dats2 = d21 + "." + m21 + "." + Convert.ToString(y2);

                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "UPDATE Журнал2 SET IDжурнала=@a1, Работник=@a2, Должность=@a3, Подразделение2=@a4, Место=@a5, Дата=@a6, Время=@a7, Обстоятельства=@a8, " +
                        " Причины=@a9, МедпунктДата=@a10, МедпунктВремя=@a11, Диагноз=@a12, ПринятыеМеры=@a13, Последствия=@a14, Инструктор=@a15 " +
                        " WHERE IDстроки = @a16";
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", 2);
                SqlCom1.Parameters.AddWithValue("@a2", textBox1.Text);
                SqlCom1.Parameters.AddWithValue("@a3", Convert.ToString(comboBox1.SelectedItem));
                SqlCom1.Parameters.AddWithValue("@a4", Convert.ToString(comboBox3.SelectedItem));
                SqlCom1.Parameters.AddWithValue("@a5", textBox11.Text);
                SqlCom1.Parameters.AddWithValue("@a6", dats);
                SqlCom1.Parameters.AddWithValue("@a7", textBox3.Text);
                SqlCom1.Parameters.AddWithValue("@a8", textBox2.Text);
                SqlCom1.Parameters.AddWithValue("@a9", textBox4.Text);
                SqlCom1.Parameters.AddWithValue("@a10", dats2);
                SqlCom1.Parameters.AddWithValue("@a11", textBox5.Text);
                SqlCom1.Parameters.AddWithValue("@a12", textBox7.Text);
                SqlCom1.Parameters.AddWithValue("@a13", textBox8.Text);
                SqlCom1.Parameters.AddWithValue("@a14", textBox9.Text);
                SqlCom1.Parameters.AddWithValue("@a15", textBox10.Text);
                SqlCom1.Parameters.AddWithValue("@a16", textBox6.Text);

                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Запись изменена.", "ИЗМЕНЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("ПУСТЫЕ ПОЛЯ НЕ ДОПУСТИМЫ!", "КОНТРОЛЬ ДАННЫХ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ShowList();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ClearAll();
        }
    }
}
