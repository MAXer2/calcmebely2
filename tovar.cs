using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Сайтец
{
    public partial class tovar : Form
    {
        public tovar(string id)
        {
            InitializeComponent();
            List<string> infaOTovare = Form1.Select("SELECT Title, Price, Image FROM `Products`" +
           " WHERE id = '" + id + "'");
            label5.Text = infaOTovare[0].ToString();
            label1.Text = infaOTovare[1].ToString() + " Р";
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void bunifuTileButton1_Click(object sender, EventArgs e)
        {
        }

        private void tovar_Load(object sender, EventArgs e)
        {
            //panel3.Location = new Point(Width / 10, Height / 10);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
