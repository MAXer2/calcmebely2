﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Label1_Click(object sender, EventArgs e)
        {
           bezPodyomaRadioButton.Checked = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            matrasComboBox.TabIndex = 1;
            bunifuFlatButton1.Visible = false;
            bunifuFlatButton2.IconZoom = 70;
        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }
        #region Функции, упрощающие запись
        /// <summary>
        /// Округление с точностью до 10 рублей
        /// </summary>
        int Round10(double price)
        {
            return (int)(price / 10 + 0.99999) * 10;
        }

        /// <summary>
        /// Убираем слово "рублей" из цены
        /// </summary>
        string ReplaceRubles(String price)
        {
            return price.Replace("(", "").Replace(" рублей)", "");
        }
        #endregion



        #region Доставка

        /// <summary>
        /// Считаем стоимость доставки
        /// </summary>
        void CalculateShippingCost()
        {
            try
            {
                int weightMatras = (int)(40 * Convert.ToDouble(matrasComboBox.Text));
                double weightFurniture = Convert.ToDouble(weightTextBox.Text);
                int priceShippingPerKg = 2;
                int pricePadik = Round10(priceShippingPerKg * (weightMatras + weightFurniture));
                if (pricePadik < 400)
                {
                    pricePadik = 400;
                }

                int priceEdinolichnik = Round10(priceShippingPerKg * (weightMatras + weightFurniture));
                if (priceEdinolichnik < 3000)
                {
                    priceEdinolichnik = 3000;
                }

                int rasst = Convert.ToInt32(rasstTextBox.Text);

                sovmDostRadioButton.Text = "(" + pricePadik.ToString() + " рублей)";
                individDostRadioButton.Text = "(" + priceEdinolichnik.ToString() + " рублей)";
                mezhgorodRadioButton.Text = "(" + (pricePadik + 30 * rasst).ToString() + " рублей)";
            }
            catch (Exception) { }

            CalculateCost();
        }

        /// <summary>
        /// Самовывоз
        /// </summary>
        private void samovyvozDostavkaClick(object sender, EventArgs e)
        {
            shippingTextBox.Text = "0";
            CalculateCost();
        }

        /// <summary>
        /// Индивидуальная доставка
        /// </summary>
        private void individDostavkaClick(object sender, EventArgs e)
        {
            shippingTextBox.Text = ReplaceRubles(individDostRadioButton.Text);
            CalculateCost();
        }

        /// <summary>
        /// Выбрана совместная доставка
        /// </summary>
        private void sovmDostavkaClick(object sender, EventArgs e)
        {
            shippingTextBox.Text = ReplaceRubles(sovmDostRadioButton.Text);
            CalculateCost();
        }

        /// <summary>
        /// Межгород
        /// </summary>
        private void mezhgorodDostavkaClick(object sender, EventArgs e)
        {
            shippingTextBox.Text = ReplaceRubles(mezhgorodRadioButton.Text);
            CalculateCost();
        }


        #endregion

        private void FurnitureTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                TextBox tb = (TextBox)sender;
                tb.Text = (Convert.ToInt32(tb.Text)).ToString();
            }
            catch (Exception) { }

            CalculateComplectCost();
            CalculatePodiomCost();
            CalculateShippingCost();
        }

        /// <summary>
        /// Считаем итоговую стоимость
        /// </summary>
        void CalculateCost()
        {
            try
            {
                totalTextBox.Text = "Итого: " + (
                    Convert.ToInt32(furnitureTextBox.Text) +
                    Convert.ToInt32(shippingTextBox.Text) +
                    Convert.ToInt32(podyomTextBox.Text) +
                    Convert.ToInt32(sborkaTextBox.Text)).ToString();
            }
            catch (Exception) { }
        }

        #region Подъем
        private void bezPodyomaRadioButton_Click(object sender, EventArgs e)
        {
            podyomTextBox.Text = "0";
            CalculateCost();
        }
        private void BezLiftaRadiobutton_Click(object sender, EventArgs e)
        {
            podyomTextBox.Text = ReplaceRubles(bezLiftaRadioButton.Text);
            CalculateCost();
        }

        private void sLiftomRadioButton_Click(object sender, EventArgs e)
        {
            podyomTextBox.Text = ReplaceRubles(sLiftomRadioButton.Text);
            CalculateCost();
        }

        /// <summary>
        /// Стоимость подъема
        /// </summary>
        void CalculatePodiomCost()
        {
            try
            {
                int etaj = Convert.ToInt32(etajTextBox.Text);

                int weightMatras = (int)(40 * Convert.ToDouble(matrasComboBox.Text));
                double weightFurniture = Convert.ToDouble(weightTextBox.Text);
                int pricePodyomPerKg = 1;
                int pricePod = Round10(pricePodyomPerKg * etaj * (2 * weightMatras + weightFurniture));
                if (pricePod < 300)
                {
                    pricePod = 300;
                }

                int priceLift = Round10(3 * pricePodyomPerKg * (2 * weightMatras + weightFurniture));
                if (priceLift < 300)
                {
                    priceLift = 300;
                }

                int rasst = Convert.ToInt32(rasstTextBox.Text);

                bezLiftaRadioButton.Text = "(" + pricePod.ToString() + " рублей)";
                sLiftomRadioButton.Text = "(" + priceLift.ToString() + " рублей)";
            }
            catch (Exception) { }

            CalculateCost();

        }
        #endregion

        #region Сборка

        private void bezSborkiRadioButtonClick(object sender, EventArgs e)
        {
            complectTextBox.Text = "0";
            CalculateCost();
        }
        private void Sborka2RadioButtonClick(object sender, EventArgs e)
        {
            complectTextBox.Text = ReplaceRubles(Sborka1RadioButton.Text);
            CalculateCost();
        }
        private void Sborka1RadioButtonClick(object sender, EventArgs e)
        {
            complectTextBox.Text = ReplaceRubles(Sborka2RadioButton.Text);
            CalculateCost();
        }

        /// <summary>
        /// Стоимость сборки
        /// </summary>
        void CalculateComplectCost()
        {
            try
            {
                int etaj = Convert.ToInt32(etajTextBox.Text);
                int rasst = Convert.ToInt32(rasstTextBox.Text);
                int kreplenie = Convert.ToInt32(kreplenieBox.Text);
                int weightMatras = (int)(40 * Convert.ToDouble(matrasComboBox.Text));
                double weightFurniture = Convert.ToDouble(weightTextBox.Text);

                int price6Pr = Round10(Convert.ToInt32(furnitureTextBox.Text) * 0.06);
                int priceSborkaCena = Round10(Convert.ToInt32(furnitureTextBox.Text) * 0.10);
                int priceSborkaVes10 = Round10(Convert.ToInt32(weightTextBox.Text) * 15);

                int priceSborkaVes = (priceSborkaCena > priceSborkaVes10) ? priceSborkaVes10 : priceSborkaCena;
                int priceSborkaUlyanovsk = (priceSborkaVes > price6Pr) ? priceSborkaVes : price6Pr;

                int priceSborkaNeUlyanovsk = Round10(Convert.ToInt32(priceSborkaUlyanovsk) + 22 * rasst);

                if (mezhgorodRadioButton.Checked)
                {
                    Sborka1RadioButton.Text = "(" + priceSborkaNeUlyanovsk.ToString() + " рублей)";
                    Sborka2RadioButton.Text = "(" + (priceSborkaNeUlyanovsk + 250 * kreplenie).ToString() + " рублей)";
                }
                else
                {
                    Sborka1RadioButton.Text = "(" + priceSborkaUlyanovsk.ToString() + " рублей)";
                    Sborka2RadioButton.Text = "(" + (priceSborkaUlyanovsk + 250 * kreplenie).ToString() + " рублей)";
                }
            }
            catch (Exception) { }

            CalculateCost();
        }
        #endregion

        private void RasstTextBox_TextChanged(object sender, EventArgs e)
        {
            CalculatePodiomCost();
            CalculateComplectCost();
            CalculateShippingCost();
        }

        private void MatrasTextBox_TextChanged(object sender, EventArgs e)
        {
            CalculatePodiomCost();
            CalculateShippingCost();
        }

        private void WeightTextBox_TextChanged(object sender, EventArgs e)
        {
            CalculatePodiomCost();
            CalculateShippingCost();
        }

        private void Label6_Click(object sender, EventArgs e)
        {

        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void TextBox9_TextChanged(object sender, EventArgs e)
        {
            CalculatePodiomCost();

        }

        private void BunifuDatepicker1_onValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged_1(object sender, EventArgs e)
        {
            CalculatePodiomCost();
        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Label18_Click(object sender, EventArgs e)
        {

        }

        private void PodyomTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void BunifuFlatButton1_Click(object sender, EventArgs e)
        {
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            CalculatePodiomCost();
            CalculateComplectCost();
            CalculateShippingCost();

            //tabControl1.SelectedIndex = tabControl1.SelectedIndex + 1;
          
            nextButton.Visible = true;
            prevButton.Visible = true;
           // if (tabControl1.SelectedIndex == 3)
            //{
            //    nextButton.Visible = false;
           // }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            CalculatePodiomCost();
            CalculateComplectCost();
            CalculateShippingCost();

            //tabControl1.SelectedIndex = tabControl1.SelectedIndex - 1;
            nextButton.Visible = true;
            prevButton.Visible = true;
           // if (tabControl1.SelectedIndex == 0)
          //  {
           //     prevButton.Visible = false;
           // }
        }

        Word.Paragraph addParagraph(object oRng, Word._Document oDoc, object oEndOfDoc)
        {
            Word.Paragraph oPara5;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara5 = oDoc.Content.Paragraphs.Add(ref oRng);

            return oPara5;
        }
        void zamena(Object missing, Object wrap, Object replace, Word.Find find, string from, string naChtoMenyaem)
        {
            find.Text = from;
            find.Replacement.Text = naChtoMenyaem.Replace("(", "").Replace(")", "");
            find.Execute(FindText: Type.Missing,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: wrap,
                Format: false,
                ReplaceWith: missing, Replace: replace);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Word.Application app = new Word.Application();
            Object fileName = Application.StartupPath + "\\Договор создания дополнительных услуг.docx";
            Object missing = Type.Missing;
            app.Documents.Open(ref fileName);
            Word.Find find = app.Selection.Find;
            Object wrap = Word.WdFindWrap.wdFindContinue;
            Object replace = Word.WdReplace.wdReplaceAll;
            find.Execute(FindText: Type.Missing,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: wrap,
                Format: false,
                ReplaceWith: missing, Replace: replace);
#region Замена 

            zamena(missing, wrap, replace, find, "*Здесь адрес клиента*", textBox6.Text);

            zamena(missing, wrap, replace, find, "*Здесь имя заказчика*", textBox4.Text);

            zamena(missing, wrap, replace, find, "*Здесь телефон клиента*", textBox5.Text);

            zamena(missing, wrap, replace, find, "*номер договора*", textBox3.Text);

            zamena(missing, wrap, replace, find, "*вес*", weightTextBox.Text);

            zamena(missing, wrap, replace, find, "*здесь этаж*", etajTextBox.Text);

            zamena(missing, wrap, replace, find, "*кматр*", matrasComboBox.Text);

            zamena(missing, wrap, replace, find, "*Здесь дата*", dateTimePicker1.Text);

            zamena(missing, wrap, replace, find, "*цена с доставки*", sovmDostRadioButton.Text);

            zamena(missing, wrap, replace, find, "*цена инд доставки*", individDostRadioButton.Text);

            zamena(missing, wrap, replace, find, "*цена доставки за пред*", mezhgorodRadioButton.Text);

            zamena(missing, wrap, replace, find, "*цена подъем без лифта*", bezLiftaRadioButton.Text);

            zamena(missing, wrap, replace, find, "*цена занос с лифтом*", sLiftomRadioButton.Text);

            zamena(missing, wrap, replace, find, "*цена сборки*", Sborka1RadioButton.Text);

            zamena(missing, wrap, replace, find, "*цена сборки за пред*", Sborka1RadioButton.Text);

            zamena(missing, wrap, replace, find, "*цена креления*", Sborka2RadioButton.Text);

            zamena(missing, wrap, replace, find, "*стоимость*", totalTextBox.Text);
            
#endregion
            app.ActiveDocument.SaveAsQuickStyleSet(Application.StartupPath + "\\Договор 1.docx");
            //app.ActiveDocument.Close();
            //app.Quit();
            app.Visible = true;
        }

        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {
            
        }

        private void bunifuThinButton21_Click_1(object sender, EventArgs e)
        {
            
        }

        private void totalTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {
            radioButton1.Checked = true;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            CalculatePodiomCost();
            CalculateComplectCost();
            CalculateShippingCost();
            if ((textBox3.Text.Equals("")) || (textBox4.Text.Equals("")) || (textBox5.Text.Equals("")) || (textBox6.Text.Equals("")) || (furnitureTextBox.Text.Equals("")) || (weightTextBox.Text.Equals("")) || (etajTextBox.Text.Equals("")) || (matrasComboBox.TabIndex.Equals(""))) { 
                MessageBox.Show("Вы не ввели все необходимые данные!!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }  else {

                panel7.Visible = false;
                panel6.Visible = true;
                //tabControl1.SelectedIndex = tabControl1.SelectedIndex + 1;
                nextButton.Visible = true;
            prevButton.Visible = true;

                if ( panel6.Visible == true) {
               
                    nextButton.Visible = false;
                    bunifuFlatButton1.Visible = true;
                }
           // if (tabControl1.SelectedIndex == 3)
          //  {
           //     nextButton.Visible = false;
           // }
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            CalculatePodiomCost();
            CalculateComplectCost();
            CalculateShippingCost();

            panel6.Visible = false;
            panel7.Visible = true;
            //  tabControl1.SelectedIndex = tabControl1.SelectedIndex - 1;
            nextButton.Visible = true;
          prevButton.Visible = true;
           // if (tabControl1.SelectedIndex == 0)
           // {
           //     prevButton.Visible = false;
          //  }
        }

        private void TabPage1_Click(object sender, EventArgs e)
        {

        }

        private void TabPage2_Click(object sender, EventArgs e)
        {

        }

        private void Label17_Click_1(object sender, EventArgs e)
        {
            Sborka1RadioButton.Checked = true;
        }

        private void TextBox12_TextChanged(object sender, EventArgs e)
        {
            return;
        }

        private void Panel5_Paint(object sender, PaintEventArgs e)
        {

        }

        private void TextBox4_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void BezLiftaRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            podyomTextBox.Text = ReplaceRubles(bezLiftaRadioButton.Text);
            CalculateCost();
        }

        private void SLiftomRadioButton_CheckedChanged(object sender, EventArgs e)
        {

            podyomTextBox.Text = ReplaceRubles(sLiftomRadioButton.Text);  
            CalculateCost();
        }

        private void Sborka1RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            sborkaTextBox.Text = ReplaceRubles(Sborka1RadioButton.Text);
            CalculateCost();
        }

        private void Sborka2RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            sborkaTextBox.Text = ReplaceRubles(Sborka2RadioButton.Text);
             CalculateCost();
        }

        private void BezPodyomaRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            shippingTextBox.Text = ReplaceRubles(bezPodyomaRadioButton.Text);
            CalculateCost();
        }

        private void BezSborkiRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            sborkaTextBox.Text = ReplaceRubles(bezSborkiRadioButton.Text);
            CalculateCost();            
        }

        private void Label16_Click(object sender, EventArgs e)
        {
            bezLiftaRadioButton.Checked = true;
        }

        private void Label15_Click(object sender, EventArgs e)
        {
            sLiftomRadioButton.Checked = true;
        }

        private void Label12_Click(object sender, EventArgs e)
        {
            sovmDostRadioButton.Checked = true;
        }

        private void Label13_Click(object sender, EventArgs e)
        {
            individDostRadioButton.Checked = true;
        }

        private void Label14_Click(object sender, EventArgs e)
        {
            mezhgorodRadioButton.Checked = true;
        }

        private void Label18_Click_1(object sender, EventArgs e)
        {
            bezSborkiRadioButton.Checked = true;
        }

        private void Label19_Click(object sender, EventArgs e)
        {
            Sborka2RadioButton.Checked = true;
        }

        private void TextBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void MatrasComboBox_onItemSelected(object sender, EventArgs e)
        {

        }

        private void BunifuFlatButton1_Click_1(object sender, EventArgs e)
        {

        }

        private void Panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void FurnitureTextBox_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void Label26_Click(object sender, EventArgs e)
        {

        }

        private void Label21_Click_1(object sender, EventArgs e)
        {

        }

        private void BunifuFlatButton2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FurnitureTextBox_TextChanged_1(object sender, EventArgs e)
        {
            CalculateCost();

        }

        private void WeightTextBox_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }
    }
    }

