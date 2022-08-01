using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Text;


namespace DashBoard
{
    public partial class addbriefing : Form
    {
        public addbriefing()
        {
            InitializeComponent();
        }
        string italic = "";
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            txtbriefing.SelectedText = txtbriefing.SelectedText.ToUpper();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
           
            if (!txtbriefing.SelectionFont.Italic)
            {
                italic = "On";
                txtbriefing.SelectionFont = new Font(txtbriefing.SelectionFont, FontStyle.Italic | txtbriefing.SelectionFont.Style);
            }
            else
            {
                txtbriefing.SelectionFont = new Font(txtbriefing.SelectionFont, FontStyle.Regular | txtbriefing.SelectionFont.Style);
                italic = "OFf";
            }
	{

	}
        }

        private void addbriefing_Load(object sender, EventArgs e)
        {
            InstalledFontCollection InsFonts = new InstalledFontCollection();

            foreach (FontFamily item in InsFonts.Families)
            {
                comboFont.Items.Add(item.Name);
            }
            comboFont.SelectedIndex = 0;




        }

        private void button1_Click(object sender, EventArgs e)
        {
            ColorDialog cl = new ColorDialog();

            

            if (cl.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtbriefing.SelectionColor = cl.Color;
            }

        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            txtbriefing.SelectionAlignment = HorizontalAlignment.Center;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            txtbriefing.SelectionAlignment = HorizontalAlignment.Right;
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            txtbriefing.SelectionAlignment = HorizontalAlignment.Left;
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }
}
