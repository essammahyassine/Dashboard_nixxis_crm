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
    public partial class Production : Form
    {
        public Production()
        {
            InitializeComponent();
        }

        int togmove;
        int Mvalx;
        int Mvaly;
        SqlConnection cnx;
        SqlCommand cmd;
        SqlDataAdapter adp;
        SqlDataAdapter adp2;
        DataTable dtcount;
        DataTable dtcampagn;
        DataTable dtagent;
        DataTable dthour;
        DataTable dthistorique;
        DataTable dthistorique_agent;
        DataTable dtctprojet;
        DataTable dtctagent;
        DataTable dtchart;
        DataTable dtpos;
        DataTable dtneg;
        DataTable dtca;
        DataTable dttotalca;
        DataTable dttotalbd;
        DataTable dthour_camp;
        DataTable dthour_cp;
        DataTable dthour_neg_cp;
        int total_bd;
        float total_ca;
        int test=0;
        int test2 = 0;
        int test3 = 0;
        int a = 0;

        public static string cp;
        public static string datestring;
        public static string datestring2;

        public void count_pos()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + cp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select * from SystemData s 
            left outer join Default_admin.dbo.Agents a on s.LastHandler=a.Id
            join Default_admin.dbo.Qualifications q on s.LastQualification = q.Id
            where  CONVERT(VARCHAR(10), s.LastHandlingTime, 103)='" + datestring2 + "' and s.LastQualificationPositive='1'");
            adp = new SqlDataAdapter(cmd);
            dtpos = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtpos);
            cmd.Connection.Close();
            

        }

        public void CA_BD()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + cp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select * from SystemData s 
            left outer join Default_admin.dbo.Agents a on s.LastHandler=a.Id
            join Default_admin.dbo.Qualifications q on s.LastQualification = q.Id
            where s.LastQualificationArgued=1");
            adp = new SqlDataAdapter(cmd);
            dttotalca = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dttotalca);
            cmd.Connection.Close();
            

        }


        public void total_BD()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + cp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select * from Data");
            adp = new SqlDataAdapter(cmd);
            dttotalbd = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dttotalbd);
            cmd.Connection.Close();
            
            
        }

        public void count_neg()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + cp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select * from SystemData s 
            left outer join Default_admin.dbo.Agents a on s.LastHandler=a.Id
            join Default_admin.dbo.Qualifications q on s.LastQualification = q.Id
            where  CONVERT(VARCHAR(10), s.LastHandlingTime, 103)='" + datestring2 + "' and s.LastQualificationPositive=-1");
            adp = new SqlDataAdapter(cmd);
            dtneg = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtneg);
            cmd.Connection.Close();
            

        }


        public void count_ca()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + cp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select * from SystemData
            where  CONVERT(VARCHAR(10), LastHandlingTime, 103)='" + datestring2 + "' and LastQualificationArgued=1");
            adp = new SqlDataAdapter(cmd);
            dtca = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtca);
            cmd.Connection.Close();
            

        }



        void box_positive_agent()
        {
            positive_agent();

            flowLayoutPanel2.Controls.Clear();

            Label[] MyLabel3 = new Label[dtagent.Rows.Count];

            for (int j = 0; j < dtagent.Rows.Count; j++)
            {

                Label lba = new Label();
                lba.Name = dtagent.Rows[j][0].ToString();
                lba.Text = "#  " + dtagent.Rows[j][0].ToString() + "  :  " + dtagent.Rows[j][1].ToString() + " ( " + dtagent.Rows[j][2].ToString() + " )";
                lba.AutoSize = true;
                lba.TextAlign = ContentAlignment.MiddleCenter;
                lba.Top = 10 + (22 + j);
                lba.Width = 55;
                lba.Font = new Font("Segoe UI", 10, FontStyle.Regular);
                lba.BorderStyle = 0;
                lba.TabStop = false;
                MyLabel3[j] = lba;

            }

            foreach (Control lab1 in MyLabel3)
            {
                flowLayoutPanel2.Controls.Add(lab1);
                flowLayoutPanel2.SetFlowBreak(lab1, true);
            }
        
        }

        void box_positive_campagne()
        {

            lblpositive.Text = count();

            hour_prod();

            positive_campagn();
            Negative_campagn();

            //flowLayoutPanel1.Controls.Clear();

            //Label[] MyLabel2 = new Label[dtcampagn.Rows.Count];

            //for (int i = 0; i < dtcampagn.Rows.Count; i++)
            //{

            //    Label lb = new Label();
            //    lb.Name = dtcampagn.Rows[i][0].ToString();
            //    lb.Text = "#  " + dtcampagn.Rows[i][0].ToString() + "  :  " + dtcampagn.Rows[i][1].ToString();
            //    lb.AutoSize = true;
            //    lb.TextAlign = ContentAlignment.MiddleCenter;
            //    lb.Top = 10 + (22 + i);
            //    lb.Width = 55;
            //    lb.Font = new Font("Segoe UI", 10, FontStyle.Regular);
            //    lb.BorderStyle = 0;
            //    lb.TabStop = false;
            //    MyLabel2[i] = lb;

            //}

            //foreach (Control lab in MyLabel2)
            //{
            //    flowLayoutPanel1.Controls.Add(lab);
            //    flowLayoutPanel1.SetFlowBreak(lab, true);
            //}

            try
            {
                lblhour.Text = "/ " + float.Parse(dthour.Rows[0][0].ToString()).ToString("0.00") + " ( h )";
                lblhour_camp.Text = float.Parse(dthour.Rows[0][0].ToString()).ToString("0.00");
            }
            catch (Exception)
            {
                lblhour.Text = "/ 0.00 ( h )";
                lblhour_camp.Text = "/ 0.00 ( h )";
            }
            
            
           
           

        
        }

        //--------------------------------------Nombre positive global--------------------------//
        public string count()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select * from OUT_Contact where LocalDateGroup='"+ date.SelectionStart.ToString("yyyyMMdd") +"' and Positive='1'");
            adp = new SqlDataAdapter(cmd);
            dtcount = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtcount);
            cmd.Connection.Close();
            return dtcount.Rows.Count.ToString();

        }


        //---------------------------------------Nombre positive par campagne-------------------------//
     

        //--------------------------------------------- # positive campagne-----------------------------//
        public string positive_campagn()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select count(Positive) as positive,DATEPART(HOUR, LocalDateTime) as Hour from OUT_Contact
            where Positive=1 and  CONVERT(VARCHAR(10),LocalDateTime, 3)='"+ date.SelectionStart.ToString("dd/MM/yy")+"' group by DATEPART(HOUR, LocalDateTime),Positive"  );
            adp = new SqlDataAdapter(cmd);
            dtcampagn = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtcampagn);
            cmd.Connection.Close();

            chart2.DataSource = dtcampagn;
            chart2.Series.First().XValueMember = "Hour";
            chart2.Series.First().YValueMembers = "positive";
            chart2.Series[0]["PieLabelStyle"] = "Disabled";
            chart2.ChartAreas[0].AxisX.Title = "Heur";
            chart2.ChartAreas[0].AxisY.Title = "Positive";

            return dtcampagn.Rows.Count.ToString();

        }

        public string Negative_campagn()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select count(Positive) as Negative,DATEPART(HOUR, LocalDateTime) as Hour from OUT_Contact
            where Positive=-1 and  CONVERT(VARCHAR(10),LocalDateTime, 3)='" + date.SelectionStart.ToString("dd/MM/yy") + "' group by DATEPART(HOUR, LocalDateTime),Positive");
            adp = new SqlDataAdapter(cmd);
            dtcampagn = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtcampagn);
            cmd.Connection.Close();

            chart3.DataSource = dtcampagn;
            chart3.Series.First().XValueMember = "Hour";
            chart3.Series.First().YValueMembers = "Negative";
            chart3.Series[0]["PieLabelStyle"] = "Disabled";
            chart3.ChartAreas[0].AxisX.Title = "Heur";
            chart3.ChartAreas[0].AxisY.Title = "Negative";

            return dtcampagn.Rows.Count.ToString();

        }



        public void positive_Hour_cp()
        {


            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + cp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select count(LastQualificationPositive) as positive,DATEPART(HOUR, LastHandlingTime) as Hour from SystemData
            where LastQualificationPositive=1 and  CONVERT(VARCHAR(10),dbo.SystemData.LastHandlingTime, 3)='" + date.SelectionStart.ToString("dd/MM/yy") + "' group by DATEPART(HOUR, LastHandlingTime),LastQualificationPositive");
            adp2 = new SqlDataAdapter(cmd);
            dthour_cp = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp2.Fill(dthour_cp);
            cmd.Connection.Close();

        }



        public void negative_Hour_cp()
        {

            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + cp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select count(LastQualificationPositive) as Negative,DATEPART(HOUR, LastHandlingTime) as Hour from SystemData
            where LastQualificationPositive=-1 and  CONVERT(VARCHAR(10),dbo.SystemData.LastHandlingTime, 3)='" + date.SelectionStart.ToString("dd/MM/yy") + "' group by DATEPART(HOUR, LastHandlingTime),LastQualificationPositive");
            adp2 = new SqlDataAdapter(cmd);
            dthour_neg_cp = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp2.Fill(dthour_neg_cp);
            cmd.Connection.Close();
        }
        //---------------------------------------------- # positive agent------------------------------//
        public string positive_agent()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select a.FirstName,count(positive),c.Description from OUT_Contact o 
            left outer join Default_admin.dbo.Agents a on o.OrigQualOriginatorId=a.Id
            join Default_admin.dbo.Campaigns c on o.CampaignId=c.Id
            where LocalDateGroup='" + date.SelectionStart.ToString("yyyyMMdd") + "' and Positive='1' group by a.FirstName,Positive,c.Description order by a.FirstName asc");
            adp = new SqlDataAdapter(cmd);
            dtagent = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtagent);
            cmd.Connection.Close();
            return dtagent.Rows.Count.ToString();

        }

        //---------------------------------------------- # positive agent par campagne------------------------------//
        public string positive_agent_camp()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select a.FirstName,count(positive),c.Description from OUT_Contact o 
            left outer join Default_admin.dbo.Agents a on o.OrigQualOriginatorId=a.Id
            join Default_admin.dbo.Campaigns c on o.CampaignId=c.Id
            where LocalDateGroup='" + date.SelectionStart.ToString("yyyyMMdd") + "' and Positive='1' and o.CampaignId='" + cp + "' group by a.FirstName,Positive,c.Description order by a.FirstName asc");
            adp = new SqlDataAdapter(cmd);
            dtagent = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtagent);
            cmd.Connection.Close();
            return dtagent.Rows.Count.ToString();

        }


         void hour_prod()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select CONVERT(DECIMAL(10,1),cast(sum(o.Duration)as float)/3600) as hour from OUT_Contact o 
            left outer join Default_admin.dbo.Agents a on o.OrigQualOriginatorId=a.Id
            join Default_admin.dbo.Campaigns c on o.CampaignId=c.Id
            where LocalDateGroup='" + date.SelectionStart.ToString("yyyyMMdd") + "' and o.Positive in ('1','-1')");
            adp = new SqlDataAdapter(cmd);
            dthour = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dthour);
            cmd.Connection.Close();
           

        }




         void hour_prod_camp()
         {
             cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
             cmd = new SqlCommand(@"select CONVERT(DECIMAL(10,1),cast(sum(o.Duration)as float)/3600) as hour from OUT_Contact o 
            left outer join Default_admin.dbo.Agents a on o.OrigQualOriginatorId=a.Id
            join Default_admin.dbo.Campaigns c on o.CampaignId=c.Id
            where LocalDateGroup='" + date.SelectionStart.ToString("yyyyMMdd") + "' and o.Positive in ('1','-1','0') and o.CampaignId='" + cp + "' ");
             adp = new SqlDataAdapter(cmd);
             dthour = new DataTable();
             cmd.Connection = cnx;
             cmd.Connection.Open();
             adp.Fill(dthour);
             cmd.Connection.Close();
             try
             {
                 lblhour.Text = float.Parse(dthour.Rows[0][0].ToString()).ToString("0.00") + " ( h )";

                 var span2 = TimeSpan.FromHours(double.Parse(dthour.Rows[0][0].ToString()));


                 lblhour_camp.Text = string.Format("{0}:{1:00}", (int)span2.TotalHours, span2.Minutes);
             }
             catch (Exception)
             {
                 lblhour.Text = "0.00 ( h )";
                 lblhour_camp.Text = "0.00 ( h )"; 
             }
             


         }



         void hour_historique()
         {
             cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
             cmd = new SqlCommand(@"select count(o.Positive) as nb,q.Description as statut from OUT_Contact o 
             left outer join Default_admin.dbo.Agents a on o.OrigQualOriginatorId=a.Id
             join Default_admin.dbo.Campaigns c on o.CampaignId=c.Id 
             join Default_admin.dbo.Qualifications q on o.OrigContactQualificationId = q.Id
             where LocalDateGroup='" + date.SelectionStart.ToString("yyyyMMdd") + "' and c.Id='" + cp + "' group by o.Positive,q.Description order by o.Positive desc");
             adp = new SqlDataAdapter(cmd);
             dthistorique = new DataTable();
             cmd.Connection = cnx;
             cmd.Connection.Open();
             adp.Fill(dthistorique);
             cmd.Connection.Close();
           
             //for (int i = 0; i < dthistorique.Rows.Count; i++)
             //{
             //    DataRow drow = dthistorique.Rows[i];

             //    // Only row that have not been deleted
             //    if (drow.RowState != DataRowState.Deleted)
             //    {
             //        // Define the list items
             //        ListViewItem lvi = new ListViewItem();

                     
             //        lvi.SubItems.Add(drow["nb"].ToString());
             //        lvi.SubItems.Add(drow["statut"].ToString());
             //        // Add the list items to the ListView
             //        listView1.Items.Add(lvi);
             //    }
             //}

             graphe.DataSource = dthistorique;
             graphe.Series.First().XValueMember = "statut";
             graphe.Series.First().YValueMembers = "nb";
             graphe.Series[0]["PieLabelStyle"] = "Disabled";
             graphe.Series[0].LegendText = "#PERCENT/#VAL #AXISLABEL";
         


         }



         void statut_appel()
         {
             cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
             cmd = new SqlCommand(@"select a.FirstName as Agent,CONVERT(DECIMAL(10,2),cast(sum(s.DurationState)as float)/3600) as Time  from AgentStat s
             left outer join Default_admin.dbo.Agents a on a.Id=s.AgentId
             where s.LocalDateGroup='" + date.SelectionStart.ToString("yyyyMMdd") + "'  and s.AgentStateId='5' group by a.FirstName order by cast(sum(s.DurationState)as float)/3600  desc ");
             adp = new SqlDataAdapter(cmd);
             dthistorique_agent = new DataTable();
             cmd.Connection = cnx;
             cmd.Connection.Open();
             adp.Fill(dthistorique_agent);
             cmd.Connection.Close();

             for (int i = 0; i < dthistorique_agent.Rows.Count; i++)
             {
                 DataRow drow = dthistorique_agent.Rows[i];

                 var span = TimeSpan.FromHours(double.Parse(drow["Time"].ToString()));

                 // Only row that have not been deleted
                 if (drow.RowState != DataRowState.Deleted)
                 {
                     // Define the list items
                     ListViewItem lvi = new ListViewItem();

                 

                     lvi.SubItems.Add(drow["Agent"].ToString());
                     lvi.SubItems.Add(string.Format("{0}:{1:00}:{2:00}", (int)span.TotalHours, span.Minutes, span.Seconds));
                     
                     // Add the list items to the ListView
                     listView2.Items.Add(lvi);
                 }
             }


         }



         


         void projet_actif()
         {
             cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
             cmd = new SqlCommand(@"select distinct c.Description,c.Id from OUT_Contact o 
             left outer join Default_admin.dbo.Campaigns c on o.CampaignId=c.Id
             where LocalDateGroup='" + date.SelectionStart.ToString("yyyyMMdd") + "' and c.Description is not null");
             adp = new SqlDataAdapter(cmd);
             dtctprojet = new DataTable();
             cmd.Connection = cnx;
             cmd.Connection.Open();
             adp.Fill(dtctprojet);
             cmd.Connection.Close();
             //lblprojet_actif.Text = dtctprojet.Rows.Count.ToString();

         }


         void agent_actif()
         {
             cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
             cmd = new SqlCommand(@"select distinct a.FirstName,a.Id from OUT_Contact o 
             left outer join Default_admin.dbo.Agents a on o.OrigQualOriginatorId=a.Id
             where LocalDateGroup='" + date.SelectionStart.ToString("yyyyMMdd") + "' and a.FirstName  is not null");
             adp = new SqlDataAdapter(cmd);
             dtctagent = new DataTable();
             cmd.Connection = cnx;
             cmd.Connection.Open();
             adp.Fill(dtctagent);
             cmd.Connection.Close();
            

         }


         void agent_actif_camp()
         {
             cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
             cmd = new SqlCommand(@"select distinct a.FirstName,a.Id from OUT_Contact o 
             left outer join Default_admin.dbo.Agents a on o.OrigQualOriginatorId=a.Id
             where LocalDateGroup='" + date.SelectionStart.ToString("yyyyMMdd") + "' and a.FirstName  is not null and o.CampaignId='" + cp + "'");
             adp = new SqlDataAdapter(cmd);
             dtctagent = new DataTable();
             cmd.Connection = cnx;
             cmd.Connection.Open();
             adp.Fill(dtctagent);
             cmd.Connection.Close();
             lblagent_actif.Text = dtctagent.Rows.Count.ToString();
         }

       
         public void affichtype()
         {
             cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
             cmd = new SqlCommand(@"select a.FirstName as Agent,CONVERT(DECIMAL(10,2),cast(sum(o.Duration)as float)/3600) as hour from OUT_Contact o 
             left outer join Default_admin.dbo.Agents a on o.OrigQualOriginatorId=a.Id
             join Default_admin.dbo.Campaigns c on o.CampaignId=c.Id
             where o.LocalDateGroup='" + date.SelectionStart.ToString("yyyyMMdd") + "'  and o.Positive in ('1','-1') group by a.FirstName order by cast(sum(o.Duration)as float)/3600  desc ");
             adp = new SqlDataAdapter(cmd);
             dtchart = new DataTable();
             cmd.Connection = cnx;
             cmd.Connection.Open();
             adp.Fill(dtchart);
             cmd.Connection.Close();

             chart1.DataSource = dtchart;
             chart1.Series.First().XValueMember = "Agent";
             chart1.Series.First().YValueMembers = "hour";
             chart1.Series[0]["PieLabelStyle"] = "Disabled";

         }


         public void affichtype_camp()
         {
             cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
             cmd = new SqlCommand(@"select a.FirstName as Agent,CONVERT(DECIMAL(10,2),cast(sum(o.Duration)as float)/3600) as hour from OUT_Contact o 
             left outer join Default_admin.dbo.Agents a on o.OrigQualOriginatorId=a.Id
             join Default_admin.dbo.Campaigns c on o.CampaignId=c.Id
             where o.LocalDateGroup='" + date.SelectionStart.ToString("yyyyMMdd") + "'  and o.Positive in ('1','-1') and o.CampaignId='"+ cp+"' group by a.FirstName order by cast(sum(o.Duration)as float)/3600  desc ");
             adp = new SqlDataAdapter(cmd);
             dtchart = new DataTable();
             cmd.Connection = cnx;
             cmd.Connection.Open();
             adp.Fill(dtchart);
             cmd.Connection.Close();

             chart1.DataSource = dtchart;
             chart1.Series.First().XValueMember = "Agent";
             chart1.Series.First().YValueMembers = "hour";
             chart1.Series[0]["PieLabelStyle"] = "Disabled";

         }

       


        private void pictureBox4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            panel13.Visible = false;
            panel14.Visible = false;
            panel20.Visible = false;
            panel21.Visible = false;
            panel22.Visible = false;

            projet_actif();
            agent_actif();
            timer1.Start();
            timer2.Start();
            datestring = date.SelectionStart.ToString("yyyyMMdd");
            datestring2 = date.SelectionStart.ToString("dd/MM/yyyy");
            int top = 0;
            int left = 0;
            for (int i = 0; i < dtctprojet.Rows.Count; i++)
            {
               
                Button btn = new Button();
                btn.Text = dtctprojet.Rows[i][0].ToString();
                btn.Name = dtctprojet.Rows[i][1].ToString();
                btn.ForeColor = Color.White;
                btn.FlatStyle = FlatStyle.Flat;
                btn.BackColor = Color.DimGray;
                btn.Font = new Font("Segoe UI", 10, FontStyle.Regular);
                btn.FlatAppearance.BorderColor = Color.DimGray;
                btn.Left = left;
                btn.Top = top;
                btn.Width = panel1.Width;
                btn.Height = 30;
                panel1.Controls.Add(btn);
                top += btn.Height + 2;
                btn.Tag = i;
                btn.Click += new EventHandler(btn_Click);
            }

           
            statut_appel();
            box_positive_agent();
            box_positive_campagne();
            affichtype();
            timer3.Start();



        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

   

        private void button1_Click(object sender, EventArgs e)
        {
            
        }



        void btn_Click(object sender, EventArgs e)
        {
            try
            {
                cp = ((Button)sender).Name;

                lblpanagent.Text = "Positive / " + ((Button)sender).Text + " / Agent";

                listView2.Items.Clear();
                hour_historique();

                hour_prod_camp();
                agent_actif_camp();
                statut_appel();
                positive_agent_camp();
                flowLayoutPanel2.Controls.Clear();

                Label[] MyLabel3 = new Label[dtagent.Rows.Count];

                for (int j = 0; j < dtagent.Rows.Count; j++)
                {

                    Label lba = new Label();
                    lba.Name = dtagent.Rows[j][0].ToString();
                    lba.Text = "#  " + dtagent.Rows[j][0].ToString() + "  :  " + dtagent.Rows[j][1].ToString() + " ( " + dtagent.Rows[j][2].ToString() + " )";
                    lba.AutoSize = true;
                    lba.TextAlign = ContentAlignment.MiddleCenter;
                    lba.Top = 10 + (22 + j);
                    lba.Width = 55;
                    lba.Font = new Font("Segoe UI", 10, FontStyle.Regular);
                    lba.BorderStyle = 0;
                    lba.TabStop = false;
                    MyLabel3[j] = lba;


                }

                foreach (Control lab1 in MyLabel3)
                {
                    flowLayoutPanel2.Controls.Add(lab1);
                    flowLayoutPanel2.SetFlowBreak(lab1, true);
                }
                affichtype_camp();



                for (int i = 0; i < dtctagent.Rows.Count; i++)
                {
                    foreach (ListViewItem lvw in listView2.Items)
                    {

                        if (lvw.SubItems[1].Text == dtctagent.Rows[i][0].ToString())
                        {
                            lvw.BackColor = Color.LightGray;
                        }
                    }
                }


                    chart2.DataSource = null;
                    pictureBox19.Visible = true;
                    pictureBox20.Visible = true;
                    backgroundWorker1.RunWorkerAsync();
               


            }
            catch (Exception)
            {
                
                
            }
            
           
            
            

        }

        private void date_DateChanged(object sender, DateRangeEventArgs e)
        {
            foreach (var series in chart1.Series)
            {
                series.Points.Clear();
            }
            listView2.Items.Clear();
         
            datestring = date.SelectionStart.ToString("yyyyMMdd");
            datestring2 = date.SelectionStart.ToString("dd/MM/yyyy");
            projet_actif();
            panel1.Controls.Clear();

            int top = 0;
            int left = 0;
            for (int i = 0; i < dtctprojet.Rows.Count; i++)
            {

                Button btn = new Button();
                btn.Text = dtctprojet.Rows[i][0].ToString();
                btn.Name = dtctprojet.Rows[i][1].ToString();
                btn.ForeColor = Color.White;
                btn.FlatStyle = FlatStyle.Flat;
                btn.BackColor = Color.DimGray;
                btn.Font = new Font("Segoe UI", 10, FontStyle.Regular);
                btn.FlatAppearance.BorderColor = Color.DimGray;
                btn.Left = left;
                btn.Top = top;
                btn.Width = panel1.Width;
                btn.Height = 30;
                panel1.Controls.Add(btn);
                top += btn.Height + 2;
                btn.Tag = i;
                btn.Click += new EventHandler(btn_Click);
            }

            agent_actif();
            box_positive_agent();
            box_positive_campagne();
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        private void panel6_DoubleClick(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void panel6_MouseDown(object sender, MouseEventArgs e)
        {
            togmove = 1;
            Mvalx = e.X;
            Mvaly = e.Y;
        }

        private void panel6_MouseUp(object sender, MouseEventArgs e)
        {
            togmove = 0;
        }

        private void panel6_MouseMove(object sender, MouseEventArgs e)
        {
            if (togmove == 1)
            {
                this.SetDesktopLocation(MousePosition.X - Mvalx, MousePosition.Y - Mvaly);
            }
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            try
            {
                Infos m = new Infos();
                m.ShowDialog();
            }
            catch (Exception)
            {
                
               
            }
            
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {

            panel13.Visible = false;
            panel14.Visible = false;
            panel20.Visible = false;
            panel21.Visible = false;
            panel22.Visible = false;
          
            test = 0;
            test2 = 0;
            test3 = 0;
            projet_actif();
            agent_actif();
            timer1.Start();
            timer2.Start();

            datestring = date.SelectionStart.ToString("yyyyMMdd");
            int top = 0;
            int left = 0;
            for (int i = 0; i < dtctprojet.Rows.Count; i++)
            {

                Button btn = new Button();
                btn.Text = dtctprojet.Rows[i][0].ToString();
                btn.Name = dtctprojet.Rows[i][1].ToString();
                btn.ForeColor = Color.White;
                btn.FlatStyle = FlatStyle.Flat;
                btn.BackColor = Color.DimGray;
                btn.Font = new Font("Segoe UI", 10, FontStyle.Regular);
                btn.FlatAppearance.BorderColor = Color.DimGray;
                btn.Left = left;
                btn.Top = top;
                btn.Width = panel1.Width;
                btn.Height = 30;
                panel1.Controls.Add(btn);
                top += btn.Height + 2;
                btn.Tag = i;
                btn.Click += new EventHandler(btn_Click);
            }
            lblpanagent.Text = "Positive / Agent";
          
            listView2.Items.Clear();
           
            statut_appel();
            box_positive_agent();
            box_positive_campagne();
            affichtype();
            timer3.Start();
            panel3.Visible = true;
        }

      

        private void timer1_Tick(object sender, EventArgs e)
        {
            
          
         
        }

        private void pictureBox18_Click(object sender, EventArgs e)
        {
            panel3.Visible = true;
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void panel3_BindingContextChanged(object sender, EventArgs e)
        {

        }

        private void panel3_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {

            if (test < dtctprojet.Rows.Count)
            {
                test++;
                lblprojet_actif.Text = test.ToString();
               
               
            }
            else if (test == dtctprojet.Rows.Count)
            {
                timer1.Stop();
            }

            
            
            
        }

        private void ovalShape7_Click(object sender, EventArgs e)
        {

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            
            if (test2 < dtctagent.Rows.Count)
            {
                test2++;
                lblagent_actif.Text = test2.ToString();


            }
            else if (test2 == dtctprojet.Rows.Count)
            {
                timer2.Stop();
            }
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            if (test3 < dtcount.Rows.Count)
            {
                test3++;
                lblpositive.Text = test3.ToString();


            }
            else if (test3 == dtcount.Rows.Count)
            {
                timer3.Stop();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            positive_Hour_cp();
            negative_Hour_cp();
            count_pos();
            count_neg();
            count_ca();
            CA_BD();
            total_BD();
            

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            txtcplus.Text = dtpos.Rows.Count.ToString();
            lblpositive.Text = dtpos.Rows.Count.ToString();
            txtcmoin.Text = dtneg.Rows.Count.ToString();
            txtca.Text = dtca.Rows.Count.ToString();
            total_ca = float.Parse(lblpositive.Text);
            total_bd = int.Parse(txtca.Text);
            try
            {
                txttq.Text = ((total_ca / total_bd) * 100).ToString("0.00") + " %";
            }
            catch (Exception)
            {
                
               
            }
           
            panel3.Visible = false;
            //listView2.Items[4].BackColor = Color.Red;
            panel13.Visible = true;
            panel14.Visible = true;
            panel20.Visible = true;
            panel21.Visible = true;
            panel22.Visible = true;
            pictureBox19.Visible = false;
            chart2.DataSource = dthour_cp;
            chart2.Series.First().XValueMember = "Hour";
            chart2.Series.First().YValueMembers = "positive";
            chart2.Series[0]["PieLabelStyle"] = "Disabled";
            chart2.ChartAreas[0].AxisX.Title = "Heur";
            chart2.ChartAreas[0].AxisY.Title = "Positive";
            

            pictureBox20.Visible = false;
            chart3.DataSource = dthour_neg_cp;
            chart3.Series.First().XValueMember = "Hour";
            chart3.Series.First().YValueMembers = "Negative";
            chart3.Series[0]["PieLabelStyle"] = "Disabled";
            chart3.ChartAreas[0].AxisX.Title = "Heur";
            chart3.ChartAreas[0].AxisY.Title = "Negative";
            




        }

        private void panel2_MouseDown(object sender, MouseEventArgs e)
        {
            togmove = 1;
            Mvalx = e.X;
            Mvaly = e.Y;
        }

        private void panel2_MouseMove(object sender, MouseEventArgs e)
        {
            if (togmove == 1)
            {
                this.SetDesktopLocation(MousePosition.X - Mvalx, MousePosition.Y - Mvaly);
            }
        }

        private void panel2_MouseUp(object sender, MouseEventArgs e)
        {
            togmove = 0;
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

      
        


        


    }
}
