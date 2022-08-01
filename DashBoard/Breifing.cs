using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace DashBoard
{
    public partial class Breifing : Form
    {
        public Breifing()
        {
            InitializeComponent();
        }

        SqlConnection cnx;
        SqlCommand cmd;
        SqlDataAdapter adp;
        DataTable dtchart;
        DataTable dtchartCA;
        DataTable dtctprojet;
        DataTable dtctagent;
        DataTable dtobj;
    
        string cp="";
        string projet="";
        string cpp="";
        string ag;


        public DataTable affichtype()
        {
                
                cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_"+cpp+";User ID=sa;password=SSNixxis7");
                cmd = new SqlCommand(@"select count( LastQualificationPositive ) as Positive,(select obj from count.dbo.Objectif where id='"+cpp+"') as Objectif,CONVERT(VARCHAR(10), LastHandlingTime, 3) as Date  from SystemData "+
                "where LastHandler='" + ag.ToString() + "' and LastQualificationPositive='1' and LastHandlingTime >= '" + DateTime.Parse(datedb.Text).ToString("yyyy-MM-dd") + "' and LastHandlingTime <= '" + DateTime.Parse(datefn.Text).ToString("yyyy-MM-dd") + " 23:59:59' " +
                "group by CONVERT(VARCHAR(10), LastHandlingTime, 3) order by CONVERT(VARCHAR(10), LastHandlingTime, 3)  asc");
                adp = new SqlDataAdapter(cmd);
                dtchart = new DataTable();
                cmd.Connection = cnx;
                cmd.Connection.Open();
                adp.Fill(dtchart);
                cmd.Connection.Close();
                return dtchart;

        }


        public DataTable affichCAH()
        {

            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + cpp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select count( LastQualificationPositive )/7 as Argued,(select ch from count.dbo.Objectif where id='" + cpp + "') as CH,CONVERT(VARCHAR(10), LastHandlingTime, 3) as Date  from SystemData " +
            "where LastHandler='" + ag.ToString() + "' and LastQualificationArgued='1' and LastHandlingTime >= '" + DateTime.Parse(datedb.Text).ToString("yyyy-MM-dd") + "' and LastHandlingTime <= '" + DateTime.Parse(datefn.Text).ToString("yyyy-MM-dd") + " 23:59:59' " +
            "group by CONVERT(VARCHAR(10), LastHandlingTime, 3) order by CONVERT(VARCHAR(10), LastHandlingTime, 3)  asc");
            adp = new SqlDataAdapter(cmd);
            dtchartCA = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtchartCA);
            cmd.Connection.Close();
            return dtchartCA;

        }

        void projet_actif()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select distinct c.Description,c.Id from OUT_Contact o 
             left outer join Default_admin.dbo.Campaigns c on o.CampaignId=c.Id where c.Description is not null");
            adp = new SqlDataAdapter(cmd);
            dtctprojet = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtctprojet);
            cmd.Connection.Close();
            Comboprojet.DataSource = dtctprojet;
            Comboprojet.DisplayMember = "Description";
            Comboprojet.ValueMember = "Id";
        }


        void agent_actif(string cp)
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_"+cp+";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select distinct a.FirstName,a.Id from SystemData s 
            left outer join Default_admin.dbo.Agents a on s.LastHandler=a.Id where a.FirstName not in ('') and LastHandlingTime >= '" + DateTime.Parse(datedb.Text).ToString("yyyy-MM-dd") + "' and LastHandlingTime <= '" + DateTime.Parse(datefn.Text).ToString("yyyy-MM-dd") + " 23:59:59'");
            adp = new SqlDataAdapter(cmd);
            dtctagent = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtctagent);
            cmd.Connection.Close();
            listBox1.DataSource = dtctagent;
            listBox1.DisplayMember = "FirstName";
            listBox1.ValueMember = "Id";
           
        }


    
        private void Breifing_Load(object sender, EventArgs e)
        {
            


            //date_now = "%" + System.DateTime.Now.ToString("/MM/yy");
            //chart1.DataSource = affichtype(); ;
            //chart1.Series.First().XValueMember = "Date";
            //chart1.Series.First().YValueMembers = "Positive";
            //chart1.Series[0]["PieLabelStyle"] = "Disabled";


            projet_actif();
           
   
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void date_DateChanged(object sender, DateRangeEventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            
               
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            
        }

        void btn_Click(object sender, EventArgs e)
        {
            try
            {
               
                cp = ((Button)sender).Name;
                projet = ((Button)sender).Text;
                label1.Text = "Positive : mois / " + ((Button)sender).Text;
                backgroundWorker1.RunWorkerAsync();
            }
            catch (Exception )
            {

               
            }
           

        }

        private void backgroundWorker1_DoWork_1(object sender, DoWorkEventArgs e)
        {
            affichtype();
            affichCAH();

            
            
        }

        private void backgroundWorker1_RunWorkerCompleted_1(object sender, RunWorkerCompletedEventArgs e)
        {
            chart1.DataSource = dtchart;
            chart1.Series.First().XValueMember = "Date";
            chart1.Series.First().YValueMembers = "Positive";
            chart1.Series[0]["PieLabelStyle"] = "Disabled";

            chart1.Series["Series2"].XValueMember = "Date";
            chart1.Series["Series2"].YValueMembers = "Objectif";
            chart1.Series[1]["PieLabelStyle"] = "Disabled";


            chart2.DataSource = dtchartCA;
            chart2.Series.First().XValueMember = "Date";
            chart2.Series.First().YValueMembers = "Argued";
            chart2.Series[0]["PieLabelStyle"] = "Disabled";

            chart2.Series["Series2"].XValueMember = "Date";
            chart2.Series["Series2"].YValueMembers = "CH";
            chart2.Series[1]["PieLabelStyle"] = "Disabled";



            listView1.Items.Clear();
            for (int i = 0; i < dtchart.Rows.Count; i++)
            {
                DataRow drow = dtchart.Rows[i];

                // Only row that have not been deleted
                if (drow.RowState != DataRowState.Deleted)
                {
                    // Define the list items
                    ListViewItem lvi = new ListViewItem();


                    lvi.SubItems.Add(drow["Positive"].ToString());
                    lvi.SubItems.Add(drow["Date"].ToString());
                    // Add the list items to the ListView
                    listView1.Items.Add(lvi);
                }
            }
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            cp = "";
            
            backgroundWorker1.RunWorkerAsync();
        }

        private void panel6_DoubleClick(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
        }

        private void date_DateChanged_1(object sender, DateRangeEventArgs e)
        {
            projet_actif();
           
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void Comboprojet_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                 agent_actif(Comboprojet.SelectedValue.ToString());
            }
            catch (Exception)
            {
                
                
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            cpp = Comboprojet.SelectedValue.ToString();
            ag = listBox1.SelectedValue.ToString();
            backgroundWorker1.RunWorkerAsync();
        }

        private void pictureBox3_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {

        }
    }
}
