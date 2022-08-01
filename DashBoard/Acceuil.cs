using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DashBoard
{
    public partial class Acceuil : Form
    {
        public Acceuil()
        {
            InitializeComponent();
        }


        int togmove;
        int Mvalx;
        int Mvaly;

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Production f = new Production();
            f.Show();
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"C:\Program Files\Nom de société par défaut\KGL_Briefing\Briefing_agent.exe");
            }
            catch (Exception)
            {

                System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Nom de société par défaut\KGL_Briefing\Briefing_agent.exe");
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"C:\Program Files\Nom de société par défaut\QLKGL\Scripter Quality.exe");
            }
            catch (Exception)
            {

                System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Nom de société par défaut\QLKGL\Scripter Quality.exe");
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"C:\Program Files\Nom de société par défaut\KGLPAIE\Paiement_et_absence.exe");
            }
            catch (Exception)
            {

                System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Nom de société par défaut\KGLPAIE\Paiement_et_absence.exe");
            }
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"C:\Program Files\Nom de société par défaut\RH\Recrutement et Suivi du personnel.exe");
            }
            catch (Exception)
            {

                System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Nom de société par défaut\RH\Recrutement et Suivi du personnel.exe");
            }
            
        }

        private void Acceuil_Load(object sender, EventArgs e)
        {

        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            togmove = 1;
            Mvalx = e.X;
            Mvaly = e.Y;
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (togmove == 1)
            {
                this.SetDesktopLocation(MousePosition.X - Mvalx, MousePosition.Y - Mvaly);
            }
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            togmove = 0;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Breifing b = new Breifing();
            b.Show();
        }
    }
}
