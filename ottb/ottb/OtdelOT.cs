using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ottb
{
    public partial class OtdelOT : Form
    {
        public OtdelOT()
        {
            InitializeComponent();
        }

        private void справкаОПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Spravka f = new Spravka();
            f.Show();
        }

        private void институтыФакультетыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sinstitut f = new Sinstitut();
            f.Show();
        }

        private void кафедрыОтделыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Skafedra f = new Skafedra();
            f.Show();
        }

        private void журналРегистрацииВводногоИнструктажаПоОхранеТрудаToolStripMenuItem_Click(object sender, EventArgs e)
        {
           Jurnal1 f = new Jurnal1();
            f.Show();
        }

        private void журналУчетаМикроповреждениймикротравмРаботниковToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Jurnal2 f = new Jurnal2();
            f.Show();
        }

        private void журналРегистрацииИнструктажейПоОхранеТрудаНаРабочемМестеИЦелевогоИнструктажаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Jurnal3s f = new Jurnal3s();
            f.Show();
        }

        private void журналРегистрацииИнструктажаПоПожарнойБезопасностиНаРабочемМестеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Jurnal4s f = new Jurnal4s();
            f.Show();
        }
    }
}
