using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Сайтец
{
    public partial class Form1 : Form
    {
        List<Image> banners = new List<Image>();
        int bannerIndex = 0;
        public Form1()
        {
            InitializeComponent();
            
            banners.Add(Properties.Resources.banner1);
            banners.Add(Properties.Resources.banner2);
            banners.Add(Properties.Resources.banner3);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            bannerIndex++;
            if (bannerIndex >= banners.Count)
            {
                bannerIndex = 0;
            }
                
        }

        private void апнрапрToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        public static string CONNECTION_STRING =
            "SslMode=none;" +
            "Server=db4free.net;" +
            "database=ingenerka;" +
            "port=3306;" +
            "uid=ingenerka;" +
            "pwd=Beavis1989;" +
            "old guids=true;";
        /*"SslMode=none;" +
            "Server=localhost;" +
            "database=ingenerka;" +
            "port=3306;" +
            "uid=root;" +
            "old guids=true;";*/

        public static MySqlConnection CONN;

        /// <summary>
        /// Select-запрос
        /// </summary>
        /// <param name="sqlzapros"></param>
        /// <returns></returns>
        public static List<string> Select(string sqlzapros)
        {
            List<string> res = new List<string>();
            MySqlCommand q = new MySqlCommand(sqlzapros, Form1.CONN);
            MySqlDataReader r = q.ExecuteReader();

            while (r.Read())
            {
                for (int inc = 0; inc < r.FieldCount; inc++)
                {
                    res.Add(r[inc].ToString());
                }
            }
            r.Close();

            return res;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CONN = new MySqlConnection(CONNECTION_STRING);           
            while (CONN.State != ConnectionState.Open)
            { 
                CONN.Open();
            }

            List<string> products = Select("SELECT id FROM Products");

            int x = 50;
            int y = 290;
            for (int i = 0; i < products.Count; i = i + 1 )
            {
                ProductUserControl product = new ProductUserControl(products[i]);
                product.Location = new Point(x, y);
                product.Size = new Size(220, 150);
                Controls.Add(product);
                x = x + 250;

                if (x > this.Width - 250)
                {
                    y = y + 200;
                    x = 50;
                }
            }
        }

        private void bannerPictureBox_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void bunifuTextbox1_OnTextChange(object sender, EventArgs e)
        {

        }

        private void bunifuCards1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void bunifuFlatButton2_Click(object sender, EventArgs e)
        {
            SignInForm f = new SignInForm();
            f.ShowDialog();
        }

        private void Form1_Scroll(object sender, ScrollEventArgs e)
        {
        }

        private void bunifuFlatButton3_Click(object sender, EventArgs e)
        {
            LoginForm l = new LoginForm();
            l.ShowDialog();
        }

        private void panel9_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {
            earn e1 = new earn ();
            e1.ShowDialog ();
        }
    }
}
