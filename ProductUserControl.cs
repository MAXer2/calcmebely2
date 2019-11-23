using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Сайтец
{
    public partial class ProductUserControl : UserControl
    {
        string id;

        public ProductUserControl(string _id)
        {
            InitializeComponent();
            id = _id;
            // pictureBox1.Load("https://steammachine.ru/slider/2337648-220x126.jpg");


            List<string> infaOTovare = Form1.Select("SELECT Title, Price, Image FROM `Products`" +
                " WHERE id = '" + id + "'");
            label2.Text = infaOTovare[0].ToString();
            label1.Text = infaOTovare[1].ToString() + " Р";


        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            tovar t1 = new tovar(id);
            t1.ShowDialog();
        }

        private void bunifuCards1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
