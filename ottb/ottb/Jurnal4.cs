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
    public partial class Jurnal4 : Form
    {
        public Jurnal4()
        {
            InitializeComponent();
        }
        OleDbConnection con;    //Строка соединения с БД
        OleDbCommand SqlCom;    //Переменная для Sql запросов       
        DataTable DT;           //Таблица для хранения результатов запроса
        OleDbDataAdapter DA;    //Адаптер для заполнения таблицы после запроса      
        bool ifcon = false;     //Флаг соединения с базой данных 
        String pname = "";
        String pid = "";
        String rukf = "";
        String rukd = "";

        String dater = "";
        int ndol = 0;
        String sdol = "";
        int npid = 0;
        private void ShowList()
        {
            //Процедура вывода списка в таблицу DataGridView1           
            DT = new DataTable();  //Создаем заново таблицу               
            // Указываем строку запроса и привязываем к соединению
            if (radioButton1.Checked)
                SqlCom = new OleDbCommand("SELECT * FROM Журнал4 ORDER BY Дата DESC", con);
            if (radioButton2.Checked)
                SqlCom = new OleDbCommand("SELECT * FROM Журнал4 ORDER BY Работник", con);
            if (radioButton3.Checked)
                SqlCom = new OleDbCommand("SELECT * FROM Журнал4 ORDER BY Инструктор", con);

            SqlCom.ExecuteNonQuery();
            DA = new OleDbDataAdapter(SqlCom); //Через адаптер получаем результаты запроса
            DA.Fill(DT); // Заполняем таблицу результами
            DataGridView1.DataSource = DT;  //Привязываем DataGridView1 к источнику               
            DataGridView1.Columns[0].Visible = false; //Столбец с ID невидимый для пользователя   
            DataGridView1.Columns[10].Visible = false;
            DataGridView1.Font = new Font("Times New Roman", 12);
        }

        private void ShowP()
        {
            SqlCom = new OleDbCommand("SELECT * FROM [Временная1]", con);
            OleDbDataReader dataReaderV = SqlCom.ExecuteReader();

            while (dataReaderV.Read())  //Пока не конец виртуальной таблицы
            {

                pid = Convert.ToString(dataReaderV.GetValue(0));
                pname = Convert.ToString(dataReaderV.GetValue(1));
                rukf = Convert.ToString(dataReaderV.GetValue(2));
                rukd = Convert.ToString(dataReaderV.GetValue(3));
            }
            dataReaderV.Close();    //Закрыть объект чтения
            npid = Convert.ToInt32(pid);
            OleDbCommand SqlComP = new OleDbCommand("SELECT * FROM [Штат] WHERE IDподразделения2= " + npid + " AND Блокирован='нет'", con);
            OleDbDataReader dataReaderP = SqlComP.ExecuteReader();
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox7.Items.Clear();
            comboBox8.Items.Clear();
            while (dataReaderP.Read())
            {
                comboBox1.Items.Add(dataReaderP.GetValue(2));
                comboBox2.Items.Add(dataReaderP.GetValue(0));
                dater = Convert.ToString(dataReaderP.GetValue(3));
                comboBox7.Items.Add(dater);
                comboBox8.Items.Add(dataReaderP.GetValue(4));
            }
            dataReaderP.Close();

            SqlComP = new OleDbCommand("SELECT * FROM [ПричиныПроведения]", con);
            dataReaderP = SqlComP.ExecuteReader();
            comboBox3.Items.Clear();
            while (dataReaderP.Read())
            {
                comboBox3.Items.Add(dataReaderP.GetValue(1));
            }
            dataReaderP.Close();

            SqlComP = new OleDbCommand("SELECT * FROM [Пользователи] WHERE IDподразделения2=" + npid, con);
            dataReaderP = SqlComP.ExecuteReader();
            comboBox5.Items.Clear();
            while (dataReaderP.Read())
            {
                comboBox5.Items.Add(dataReaderP.GetValue(2));
            }
            dataReaderP.Close();
        }
        private bool IfNull()
        {
            //Функция обнаружения пустых полей
            if ((comboBox1.SelectedIndex == -1) || (comboBox3.SelectedIndex == -1) || (comboBox5.SelectedIndex == -1))
            {
                return true;   //Имеются пустые поля           
            }
            else
                return false;  //Пустые поля отсутствуют         
        }
        private void Jurnal4_Load(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=ottb.accdb");
                con.Open();     //Открыть базу данных
                ifcon = true;   //Флаг поднят. Соединение с базой данных прошло успешно.
                ShowP();
                ShowList();
                Control();
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "ОШИБКА ДОСТУПА К БАЗЕ ДАННЫХ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Jurnal4_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ifcon) con.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ShowList();
        }

        private void button6_Click(object sender, EventArgs e)
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

                //Поиск id работника
                int combo1 = comboBox1.SelectedIndex;
                comboBox2.SelectedIndex = combo1;
                comboBox7.SelectedIndex = combo1;//дата рождения
                comboBox8.SelectedIndex = combo1;//IDдолжности
                String dr = Convert.ToString(comboBox7.SelectedItem);

                DateTime tdr = Convert.ToDateTime(dr);
                int god = tdr.Year;
                int iddol = Convert.ToInt32(comboBox8.SelectedItem);
                OleDbCommand SqlComP = new OleDbCommand("SELECT Наименование FROM [Должности] WHERE IDдолжности=" + iddol, con);
                OleDbDataReader dataReaderP = SqlComP.ExecuteReader();
                while (dataReaderP.Read())
                {
                    sdol = Convert.ToString(dataReaderP.GetValue(0));
                }
                dataReaderP.Close();

                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "INSERT INTO Журнал4 (IDжурнала, IDподразделения2, Дата, IDработника, ГодРождения, IDдолжности, Причина, Инструктор, ДолжностьИнструктора, " +
                      " Работник, Должность) " +
                      " VALUES(@a1, @a2, @a3, @a4, @a5, @a6, @a7, @a8, @a9, @a10, @a11)";
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", 4);
                SqlCom1.Parameters.AddWithValue("@a2", npid);
                SqlCom1.Parameters.AddWithValue("@a3", dats);
                SqlCom1.Parameters.AddWithValue("@a4", Convert.ToString(comboBox2.SelectedItem));//IDработника
                SqlCom1.Parameters.AddWithValue("@a5", god);//Год
                SqlCom1.Parameters.AddWithValue("@a6", Convert.ToString(comboBox8.SelectedItem));//IDдолжности
                SqlCom1.Parameters.AddWithValue("@a7", Convert.ToString(comboBox3.SelectedItem));//Причина
                SqlCom1.Parameters.AddWithValue("@a8", Convert.ToString(comboBox5.SelectedItem));//Инструктор
                SqlCom1.Parameters.AddWithValue("@a9", rukd);
                SqlCom1.Parameters.AddWithValue("@a10", Convert.ToString(comboBox1.SelectedItem));
                SqlCom1.Parameters.AddWithValue("@a11", sdol);

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

                //Поиск id работника
                int combo1 = comboBox1.SelectedIndex;
                comboBox2.SelectedIndex = combo1;
                comboBox7.SelectedIndex = combo1;//дата рождения
                comboBox8.SelectedIndex = combo1;//IDдолжности
                String dr = Convert.ToString(comboBox7.SelectedItem);

                DateTime tdr = Convert.ToDateTime(dr);
                int god = tdr.Year;
                int iddol = Convert.ToInt32(comboBox8.SelectedItem);
                OleDbCommand SqlComP = new OleDbCommand("SELECT Наименование FROM [Должности] WHERE IDдолжности=" + iddol, con);
                OleDbDataReader dataReaderP = SqlComP.ExecuteReader();
                while (dataReaderP.Read())
                {
                    sdol = Convert.ToString(dataReaderP.GetValue(0));
                }
                dataReaderP.Close();

                OleDbCommand SqlCom1 = new OleDbCommand();
                SqlCom1.CommandText = "UPDATE Журнал4 SET IDжурнала=@a1, IDподразделения2=@a2, Дата=@a3, IDработника=@a4, ГодРождения=@a5, IDдолжности=@a6, Причина=@a7, Инструктор=@a8, ДолжностьИнструктора=@a9, " +
                      " Работник=@a10, Должность=@a11 WHERE IDстроки =@a12";
                SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                SqlCom1.Parameters.AddWithValue("@a1", 4);
                SqlCom1.Parameters.AddWithValue("@a2", npid);
                SqlCom1.Parameters.AddWithValue("@a3", dats);
                SqlCom1.Parameters.AddWithValue("@a4", Convert.ToString(comboBox2.SelectedItem));//IDработника
                SqlCom1.Parameters.AddWithValue("@a5", god);//Год
                SqlCom1.Parameters.AddWithValue("@a6", Convert.ToString(comboBox8.SelectedItem));//IDдолжности
                SqlCom1.Parameters.AddWithValue("@a7", Convert.ToString(comboBox3.SelectedItem));//Причина
                SqlCom1.Parameters.AddWithValue("@a8", Convert.ToString(comboBox5.SelectedItem));//Инструктор
                SqlCom1.Parameters.AddWithValue("@a9", rukd);
                SqlCom1.Parameters.AddWithValue("@a10", Convert.ToString(comboBox1.SelectedItem));
                SqlCom1.Parameters.AddWithValue("@a11", sdol);
                SqlCom1.Parameters.AddWithValue("@a12", textBox6.Text);

                SqlCom1.Connection = con;
                SqlCom1.ExecuteScalar(); //Выполняем запрос
                ShowList();
                MessageBox.Show("Запись изменена.", "ИЗМЕНЕНИЕ ЗАПИСИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("ПУСТЫЕ ПОЛЯ НЕ ДОПУСТИМЫ!", "КОНТРОЛЬ ДАННЫХ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void Control()
        {
            DateTime datt = DateTime.Today;
            String sdatt = Convert.ToString(datt).Substring(0, 10);
            DateTime dat = DateTime.Today;
            dat = dat.AddDays(-170);            
            int fl = 0;
            int[] id = new int[100];
            String[] fio = new string[100];
            DateTime[] data = new DateTime[100];
            int k = -1;
            DateTime d2 = DateTime.Today;
            String err = "ЖУРНАЛ РЕГИСТРАЦИИ ИНСТРУКТАЖА ПО ПОЖАРНОЙ БЕЗОПАСНОСТИ  НА РАБОЧЕМ МЕСТЕ\n\n";
            err = err+ "Текущая дата: " +sdatt+"\nПросрочены даты прохождения инструктажей\nили осталось менее 2 недель:\n\n";
            OleDbCommand SqlCom1 = new OleDbCommand();
            SqlCom1.CommandText = "SELECT DISTINCT IDработника, Работник FROM Журнал4 WHERE IDподразделения2 = " + pid;
            SqlCom1.Connection = con;
            OleDbDataReader dataReader1 = SqlCom1.ExecuteReader();
            while (dataReader1.Read())
            {
                k = k + 1;
                id[k] = Convert.ToInt32(dataReader1.GetValue(0));
                fio[k]= Convert.ToString(dataReader1.GetValue(1));
                fl = 1;
            }
            dataReader1.Close();
            if (fl == 1)
            {
                for (int i = 0; i <= k; i++)
                {
                    SqlCom1.CommandText = "SELECT max(Дата) FROM Журнал4 WHERE IDработника = " + id[i];
                    SqlCom1.Connection = con;
                    dataReader1 = SqlCom1.ExecuteReader();
                    while (dataReader1.Read())
                    {
                        d2 = Convert.ToDateTime(dataReader1.GetValue(0));
                    }
                    data[i] = d2;
                    dataReader1.Close();
                }
                int fl2 = 0;
                for (int i = 0; i <= k; i++)
                {
                    if (data[i] < dat)
                    {
                        fl2 = 1;
                        err = err + fio[i] + ".\nПоследняя дата прохождения инструктажа: " + Convert.ToString(data[i]).Substring(0,10) + "\n\n";
                    }
                }
                if (fl2 == 1)
                {
                    MessageBox.Show(err, "КОНТРОЛЬ СРОКОВ ПРОХОЖДЕНИЯ ИНСТРУКТАЖЕЙ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    string MyFile1 = "Журнал4.txt";
                    var MyWrite = new System.IO.StreamWriter(MyFile1, false);
                    MyWrite.WriteLine(err, true);
                    MyWrite.Close();
                    try
                    {
                        System.Diagnostics.Process.Start("Notepad", "Журнал4.txt");
                    }
                    catch
                    {
                        MessageBox.Show("Файл Журнал4 не найден!", "ОШИБКА ЧТЕНИЯ ФАЙЛА", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }                    
                }
                else
                    MessageBox.Show("Отсутствуют работники с просроченной датой инструктажа", "КОНТРОЛЬ СРОКОВ ИНСТРУКТАЖЕЙ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
                              
         
        private void button1_Click(object sender, EventArgs e)
        {
            Control();
        }
    }
}
