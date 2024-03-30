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
    public partial class Jurnal3s : Form
    {
        public Jurnal3s()
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
                SqlCom = new OleDbCommand("SELECT * FROM Журнал3 ORDER BY Дата DESC", con);
            if (radioButton2.Checked)
                SqlCom = new OleDbCommand("SELECT * FROM Журнал3 ORDER BY Работник", con);
            if (radioButton3.Checked)
                SqlCom = new OleDbCommand("SELECT * FROM Журнал3 ORDER BY Инструктор", con);

            SqlCom.ExecuteNonQuery();
            DA = new OleDbDataAdapter(SqlCom); //Через адаптер получаем результаты запроса
            DA.Fill(DT); // Заполняем таблицу результами
            DataGridView1.DataSource = DT;  //Привязываем DataGridView1 к источнику               
            DataGridView1.Columns[0].Visible = false; //Столбец с ID невидимый для пользователя   
            DataGridView1.Columns[11].Visible = false;
            DataGridView1.Columns[12].Visible = false;
            DataGridView1.Font = new Font("Times New Roman", 12);
        }

        private void Jurnal3s_Load(object sender, EventArgs e)
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

        private void button5_Click(object sender, EventArgs e)
        {
            ShowList();
        }

        private void Jurnal3s_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ifcon) con.Close();
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
            int[] podr = new int[100];
            String[] spodr = new String[100];
            DateTime[] data = new DateTime[100];
            int k = -1;
            DateTime d2 = DateTime.Today;
            String err = "ЖУРНАЛ РЕГИСТРАЦИИ ИНСТРУКТАЖЕЙ ПО ОХРАНЕ ТРУДА НА РАБОЧЕМ МЕСТЕ И ЦЕЛЕВОГО ИНСТРУКТАЖА\n\n";
            err = err + "Текущая дата: " + sdatt + "\nПросрочены даты прохождения инструктажей\nили осталось менее 2 недель:\n\n";
            OleDbCommand SqlCom1 = new OleDbCommand();
            SqlCom1.CommandText = "SELECT DISTINCT IDработника, Работник, IDподразделения2 FROM Журнал3";

            SqlCom1.Connection = con;
            OleDbDataReader dataReader1 = SqlCom1.ExecuteReader();
            while (dataReader1.Read())
            {
                k = k + 1;
                id[k] = Convert.ToInt32(dataReader1.GetValue(0));
                fio[k] = Convert.ToString(dataReader1.GetValue(1));
                podr[k] = Convert.ToInt32(dataReader1.GetValue(2));
                fl = 1;
            }
            dataReader1.Close();
            for (int i = 0; i <= k; i++)
            {
                SqlCom1.CommandText = "SELECT НаименованиеПолное FROM Подразделения2 WHERE IDподразделения2 = " + podr[i];
                SqlCom1.Connection = con;
                dataReader1 = SqlCom1.ExecuteReader();
                while (dataReader1.Read())
                {
                    spodr[i] = Convert.ToString(dataReader1.GetValue(0));
                }
                dataReader1.Close();
            }

            if (fl == 1)
            {
                for (int i = 0; i <= k; i++)
                {
                    SqlCom1.CommandText = "SELECT max(Дата) FROM Журнал3 WHERE IDработника = " + id[i];
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
                        err = err + fio[i] + " --> " + spodr[i] + "\nПоследняя дата прохождения инструктажа: " + Convert.ToString(data[i]).Substring(0, 10) + "\n\n";
                    }
                }
                if (fl2 == 1)
                {
                    MessageBox.Show(err, "КОНТРОЛЬ СРОКОВ ПРОХОЖДЕНИЯ ИНСТРУКТАЖЕЙ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    string MyFile1 = "Журнал3.txt";
                    var MyWrite = new System.IO.StreamWriter(MyFile1, false);
                    MyWrite.WriteLine(err, true);
                    MyWrite.Close();
                    try
                    {
                        System.Diagnostics.Process.Start("Notepad", "Журнал3.txt");
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
