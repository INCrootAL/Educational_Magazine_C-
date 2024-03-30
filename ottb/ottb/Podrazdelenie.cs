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
    public partial class Podrazdelenie : Form
    {
        public Podrazdelenie()
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
        private void Podrazdelenie_Load(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=ottb.accdb");
                con.Open();     //Открыть базу данных
                ifcon = true;   //Флаг поднят. Соединение с базой данных прошло успешно. 
                
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
                label1.Text = pname;
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "ОШИБКА ДОСТУПА К БАЗЕ ДАННЫХ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Podrazdelenie_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ifcon) con.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Jurnal3 f = new Jurnal3();
            f.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Shtat f = new Shtat();
            f.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Jurnal4 f = new Jurnal4();
            f.Show();
        }
    }
}
