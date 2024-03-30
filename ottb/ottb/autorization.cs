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
    public partial class autorization : Form
    {
        public autorization()
        {
            InitializeComponent();
        }
        OleDbConnection con;    //Строка соединения с БД
        OleDbCommand SqlCom;    //Переменная для Sql запросов       
        DataTable DT;           //Таблица для хранения результатов запроса
        OleDbDataAdapter DA;    //Адаптер для заполнения таблицы после запроса      
        bool ifcon = false;     //Флаг соединения с базой данных    

        private void autorization_Load(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=ottb.accdb");
                con.Open();     //Открыть базу данных
                ifcon = true;   //Флаг поднят. Соединение с базой данных прошло успешно.                   
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "ОШИБКА ДОСТУПА К БАЗЕ ДАННЫХ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int flag = 0;
            String log1 = "", parol1 = "", role = "";
            String log = "'" + Convert.ToString(textBox1.Text) + "'";
            String parol = "'" + Convert.ToString(textBox2.Text) + "'";
            String blok = "";

            String pname ="";
            String pid = "";
            String rukf = "";
            String rukd = "";
            SqlCom = new OleDbCommand("SELECT * FROM [Пользователи] WHERE Логин = " + log + " AND Пароль = " + parol, con);
            OleDbDataReader dataReaderV = SqlCom.ExecuteReader();
            while (dataReaderV.Read())  //Пока не конец виртуальной таблицы
            {
                log1 = Convert.ToString(dataReaderV.GetValue(3));
                parol1 = Convert.ToString(dataReaderV.GetValue(4));
                role = Convert.ToString(dataReaderV.GetValue(5));
                blok= Convert.ToString(dataReaderV.GetValue(6));
                pname= Convert.ToString(dataReaderV.GetValue(1));
                pid = Convert.ToString(dataReaderV.GetValue(8)); 
                rukf = Convert.ToString(dataReaderV.GetValue(2));
                rukd = Convert.ToString(dataReaderV.GetValue(7));

                flag = 1;
            }
            dataReaderV.Close();    //Закрыть объект чтения
            if (flag == 0)
            {
                MessageBox.Show("Неправильный логин или пароль!", "АВТОРИЗАЦИЯ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else

            {
                textBox1.Clear();
                textBox2.Clear();
                if (blok == "да")
                {
                    MessageBox.Show("Доступ в систему запрещен. \nОбратитесь к администратору системы.", "АВТОРИЗАЦИЯ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (role == "администратор")
                    {
                        Admin f = new Admin();
                        f.Show();
                    }
                    if (role == "работник ООТиТБ")
                    {
                        OtdelOT f = new OtdelOT();
                        f.Show();
                    }
                    if (role == "кафедра")
                    {
                        textBox3.Text = pid;
                        textBox4.Text = pname;
                        textBox5.Text = rukf;
                        textBox6.Text = rukd;

                        OleDbCommand SqlCom1 = new OleDbCommand();
                        SqlCom1.CommandText = "UPDATE [Временная1] SET id=@a1, Название=@a2, Инструктор=@a3, Должность=@a4";
                        SqlCom1.Parameters.Clear(); //Очистка параметров вызова
                        SqlCom1.Parameters.AddWithValue("@a1", textBox3.Text);
                        SqlCom1.Parameters.AddWithValue("@a2", textBox4.Text);
                        SqlCom1.Parameters.AddWithValue("@a3", textBox5.Text);
                        SqlCom1.Parameters.AddWithValue("@a4", textBox6.Text); 
                        SqlCom1.Connection = con;
                        SqlCom1.ExecuteScalar(); //Выполняем запрос
                        Podrazdelenie f = new Podrazdelenie();
                        f.Show();
                    }
                }
            }
        }

        private void autorization_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ifcon) con.Close();
        }
    }
}
