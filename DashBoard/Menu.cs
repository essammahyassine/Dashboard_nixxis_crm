using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace DashBoard
{
    public partial class Infos : Form
    {
        public Infos()
        {
            InitializeComponent();
        }

        int togmove;
        int Mvalx;
        int Mvaly;
        SqlConnection cnx;
        SqlCommand cmd;
        SqlDataAdapter adp;
        DataTable dtctagent;
        DataTable dthistorique;
        DataTable dtpos;
        DataTable dtneg;
        DataTable dtonligne;
        DataTable dtwaiting;
        DataTable dtwrap;
        DataTable dtpause;
        DateTime date;
        string id_agent;

        void total_h()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select CONVERT(DECIMAL(10,2),cast(sum(DurationState)as float)/3600)from AgentStat
            where LocalDateGroup='" + Production.datestring + "' and AgentStateId in ('5','3','4','6') and AgentId='" + id_agent + "'");
            adp = new SqlDataAdapter(cmd);
            dtonligne = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtonligne);
            cmd.Connection.Close();

            try
            {

                var span = TimeSpan.FromHours(double.Parse(dtonligne.Rows[0][0].ToString()));
                lblhour_work.Text = string.Format("{0}:{1:00}", (int)span.TotalHours, span.Minutes);
            }
            catch (Exception)
            {


            }
        }

        void onligne()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select CONVERT(DECIMAL(10,2),cast(sum(DurationState)as float)/3600)from AgentStat
            where LocalDateGroup='"+ Production.datestring +"' and AgentStateId='5' and AgentId='"+id_agent+"'");
            adp = new SqlDataAdapter(cmd);
            dtonligne = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtonligne);
            cmd.Connection.Close();

            try
            {

                var span = TimeSpan.FromHours(double.Parse(dtonligne.Rows[0][0].ToString()));
                lblonligne.Text = string.Format("{0}:{1:00}:{2:00}", (int)span.TotalHours, span.Minutes, span.Seconds);
            }
            catch (Exception)
            {
                
               
            }
        }

        void waiting()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select CONVERT(DECIMAL(10,2),cast(sum(DurationState)as float)/3600)from AgentStat
            where LocalDateGroup='" + Production.datestring + "' and AgentStateId='4' and AgentId='" + id_agent + "'");
            adp = new SqlDataAdapter(cmd);
            dtwaiting = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtwaiting);
            cmd.Connection.Close();

            try
            {

                var span = TimeSpan.FromHours(double.Parse(dtwaiting.Rows[0][0].ToString()));
                lblwaiting.Text = string.Format("{0}:{1:00}:{2:00}", (int)span.TotalHours, span.Minutes, span.Seconds);
            }
            catch (Exception)
            {
                
                
            }
        }

        void Wrapup()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select CONVERT(DECIMAL(10,2),cast(sum(DurationState)as float)/3600)from AgentStat
            where LocalDateGroup='" + Production.datestring + "' and AgentStateId='6' and AgentId='" + id_agent + "'");
            adp = new SqlDataAdapter(cmd);
            dtwrap = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtwrap);
            cmd.Connection.Close();

            try
            {

                var span = TimeSpan.FromHours(double.Parse(dtwrap.Rows[0][0].ToString()));
                lblwrap.Text = string.Format("{0}:{1:00}:{2:00}", (int)span.TotalHours, span.Minutes, span.Seconds);
            }
            catch (Exception)
            {
                
              
            }
        }

        void Pause()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select CONVERT(DECIMAL(10,2),cast(sum(DurationState)as float)/3600)from AgentStat
            where LocalDateGroup='" + Production.datestring + "' and AgentStateId='3' and AgentId='" + id_agent + "'");
            adp = new SqlDataAdapter(cmd);
            dtpause = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtpause);
            cmd.Connection.Close();

            try
            {

                var span = TimeSpan.FromHours(double.Parse(dtpause.Rows[0][0].ToString()));
                lblpause.Text = string.Format("{0}:{1:00}:{2:00}", (int)span.TotalHours, span.Minutes, span.Seconds);
            }
            catch (Exception)
            {
                
                
            }
            

           
        }


        public string count_pos()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_"+Production.cp+";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select * from SystemData s 
            left outer join Default_admin.dbo.Agents a on s.LastHandler=a.Id
            join Default_admin.dbo.Qualifications q on s.LastQualification = q.Id
            where  CONVERT(VARCHAR(10), s.LastHandlingTime, 103)='"+Production.datestring2+"' and s.LastQualificationPositive='1' and s.LastHandler='" + id_agent + "'");
            adp = new SqlDataAdapter(cmd);
            dtpos = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtpos);
            cmd.Connection.Close();
            return dtpos.Rows.Count.ToString();

        }


        public string count_neg()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + Production.cp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select * from SystemData s 
            left outer join Default_admin.dbo.Agents a on s.LastHandler=a.Id
            join Default_admin.dbo.Qualifications q on s.LastQualification = q.Id
            where  CONVERT(VARCHAR(10), s.LastHandlingTime, 103)='" + Production.datestring2 + "' and s.LastQualificationArgued=1 and s.LastHandler='" + id_agent + "'");
            adp = new SqlDataAdapter(cmd);
            dtneg = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtneg);
            cmd.Connection.Close();
            return dtneg.Rows.Count.ToString();

        }

        void agent_actif()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_ContactRouteReport;User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select distinct a.FirstName,a.Id from OUT_Contact o 
             left outer join Default_admin.dbo.Agents a on o.OrigQualOriginatorId=a.Id
             join Default_admin.dbo.Campaigns c on o.CampaignId=c.Id 
             where LocalDateGroup='" + Production.datestring + "' and a.FirstName  is not null and o.CampaignId='"+ Production.cp +"'");
            adp = new SqlDataAdapter(cmd);
            dtctagent = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dtctagent);
            cmd.Connection.Close();
            id_agent = dtctagent.Rows[0][1].ToString();
            stats_historique();
            lblpos.Text = count_pos();
            lblneg.Text = count_neg();
            Nagent.Text = dtctagent.Rows[0][0].ToString();
            lblpanstats.Text = "Positif / " + dtctagent.Rows[0][0].ToString();
            lblpanag.Text = "Négatif / " + dtctagent.Rows[0][0].ToString();
            lblpanag2.Text = "Neutre / " + dtctagent.Rows[0][0].ToString();

            pos_hist();
            neut_hist();
            neg_hist();
            total_h();
            onligne();
            waiting();
            Wrapup();
            Pause();
            chart2_f();
           

        }


        void stats_historique()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_"+ Production.cp +";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select count(s.LastQualificationPositive) as nb,q.Description as statut from SystemData s 
            left outer join Default_admin.dbo.Agents a on s.LastHandler=a.Id
            join Default_admin.dbo.Qualifications q on s.LastQualification = q.Id
            where  CONVERT(VARCHAR(10), s.LastHandlingTime, 103)='" + Production.datestring2 + "' and s.LastHandler='" + id_agent + "'  group by s.LastQualificationPositive,q.Description order by s.LastQualificationPositive desc");
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

            chart_statut.DataSource = dthistorique;
            chart_statut.Series.First().XValueMember = "statut";
            chart_statut.Series.First().YValueMembers = "nb";
            chart_statut.Series[0]["PieLabelStyle"] = "Disabled";


        }





        //------------------------------------------------------------------------------------//
        void neg_hist()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + Production.cp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select count(s.LastQualificationPositive) as nb,q.Description as statut from SystemData s 
            left outer join Default_admin.dbo.Agents a on s.LastHandler=a.Id
            join Default_admin.dbo.Qualifications q on s.LastQualification = q.Id
            where  CONVERT(VARCHAR(10), s.LastHandlingTime, 103)='" + Production.datestring2 + "' and s.LastHandler='" + id_agent + "' and  s.LastQualificationPositive='-1' group by s.LastQualificationPositive,q.Description order by s.LastQualificationPositive desc");
            adp = new SqlDataAdapter(cmd);
            dthistorique = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dthistorique);
            cmd.Connection.Close();

            for (int i = 0; i < dthistorique.Rows.Count; i++)
            {
                DataRow drow = dthistorique.Rows[i];

                // Only row that have not been deleted
                if (drow.RowState != DataRowState.Deleted)
                {
                    // Define the list items
                    ListViewItem lvi = new ListViewItem();


                    lvi.SubItems.Add(drow["nb"].ToString());
                    lvi.SubItems.Add(drow["statut"].ToString());
                    // Add the list items to the ListView
                    listneg.Items.Add(lvi);
                }
            }
        }

        void pos_hist()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + Production.cp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select count(s.LastQualificationPositive) as nb,q.Description as statut from SystemData s 
            left outer join Default_admin.dbo.Agents a on s.LastHandler=a.Id
            join Default_admin.dbo.Qualifications q on s.LastQualification = q.Id
            where  CONVERT(VARCHAR(10), s.LastHandlingTime, 103)='" + Production.datestring2 + "' and s.LastHandler='" + id_agent + "' and  s.LastQualificationPositive='1' group by s.LastQualificationPositive,q.Description order by s.LastQualificationPositive desc");
            adp = new SqlDataAdapter(cmd);
            dthistorique = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dthistorique);
            cmd.Connection.Close();

            for (int i = 0; i < dthistorique.Rows.Count; i++)
            {
                DataRow drow = dthistorique.Rows[i];

                // Only row that have not been deleted
                if (drow.RowState != DataRowState.Deleted)
                {
                    // Define the list items
                    ListViewItem lvi = new ListViewItem();


                    lvi.SubItems.Add(drow["nb"].ToString());
                    lvi.SubItems.Add(drow["statut"].ToString());
                    // Add the list items to the ListView
                    listpos.Items.Add(lvi);
                }
            }
        }

        void neut_hist()
        {
            cnx = new SqlConnection("Data Source=192.168.1.9;Initial Catalog=Default_Data_" + Production.cp + ";User ID=sa;password=SSNixxis7");
            cmd = new SqlCommand(@"select count(s.LastQualificationPositive) as nb,q.Description as statut from SystemData s 
            left outer join Default_admin.dbo.Agents a on s.LastHandler=a.Id
            join Default_admin.dbo.Qualifications q on s.LastQualification = q.Id
            where  CONVERT(VARCHAR(10), s.LastHandlingTime, 103)='" + Production.datestring2 + "' and s.LastHandler='" + id_agent + "' and  s.LastQualificationPositive='0' group by s.LastQualificationPositive,q.Description order by s.LastQualificationPositive desc");
            adp = new SqlDataAdapter(cmd);
            dthistorique = new DataTable();
            cmd.Connection = cnx;
            cmd.Connection.Open();
            adp.Fill(dthistorique);
            cmd.Connection.Close();

            for (int i = 0; i < dthistorique.Rows.Count; i++)
            {
                DataRow drow = dthistorique.Rows[i];

                // Only row that have not been deleted
                if (drow.RowState != DataRowState.Deleted)
                {
                    // Define the list items
                    ListViewItem lvi = new ListViewItem();


                    lvi.SubItems.Add(drow["nb"].ToString());
                    lvi.SubItems.Add(drow["statut"].ToString());
                    // Add the list items to the ListView
                    listneut.Items.Add(lvi);
                }
            }
        }

























        void chart2_f()
        {

            foreach (var series in chart_oligne.Series)
            {
                series.Points.Clear();
            }

            float a = float.Parse(dtonligne.Rows[0][0].ToString());
            float b = float.Parse(dtwaiting.Rows[0][0].ToString());
            float c = float.Parse(dtwrap.Rows[0][0].ToString());
            float d = float.Parse(dtpause.Rows[0][0].ToString());


            chart_oligne.DataSource = "";

            chart_oligne.Series["Onligne"].Color = Color.LimeGreen;
            chart_oligne.Series["Onligne"].Points.AddXY("", a);

            chart_oligne.Series["Waiting"].Color = Color.Orange;
            chart_oligne.Series["Waiting"].Points.AddXY("", b);


            chart_oligne.Series["WrapUp"].Points.AddXY("", c);
            chart_oligne.Series["WrapUp"].Color = Color.Yellow;


            chart_oligne.Series["Pause"].Points.AddXY("", d);
            chart_oligne.Series["Pause"].Color = Color.Red;
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Menu_Load(object sender, EventArgs e)
        {
            
            agent_actif();
           

            int top = 0;
            int left = 0;
            for (int i = 0; i < dtctagent.Rows.Count; i++)
            {

                Button btn = new Button();
                btn.Text = dtctagent.Rows[i][0].ToString();
                btn.Name = dtctagent.Rows[i][1].ToString();
                btn.ForeColor = Color.Orange;
                btn.FlatStyle = FlatStyle.Flat;
                btn.BackColor = Color.DimGray;
                btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
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


        }
        void btn_Click(object sender, EventArgs e)
        {
            id_agent = ((Button)sender).Name;
            listpos.Items.Clear();
            listneg.Items.Clear();
            listpos.Items.Clear();
            listneut.Items.Clear();
            stats_historique();
            pos_hist();
            neut_hist();
            neg_hist();
            lblpos.Text = count_pos();
            lblneg.Text = count_neg();
            try
            {
                ratiolbl.Text = (float.Parse(lblpos.Text) / int.Parse(lblneg.Text)).ToString("0.00 %");
            }
            catch (Exception)
            {

                ratiolbl.Text = "0.00 %";
            }
           
            Nagent.Text = ((Button)sender).Text;
            lblpanstats.Text = "Positif / " + ((Button)sender).Text;
            lblpanag.Text = "Négatif / " + ((Button)sender).Text;
            lblpanag2.Text = "Neutre / " + ((Button)sender).Text;
            total_h();
            onligne();
            waiting();
            Wrapup();
            Pause();
            chart2_f();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlApp.DisplayInsertOptions = false;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Range["E3", "F3"].Value = "AGENT";
            xlWorkSheet.get_Range("A:A", System.Type.Missing).EntireColumn.ColumnWidth = 10;
            xlWorkSheet.Cells[4, 1].borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            xlWorkSheet.Cells[4, 1].borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            xlWorkSheet.Cells[4, 1].borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            xlWorkSheet.Cells[4, 1].borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            xlWorkSheet.Cells[4, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            SaveFileDialog op;
            op = new SaveFileDialog();
            op.Filter = "Fichier Excel (.xls)|*.xls|Fichier CSV (.csv)|*.csv ";
            if (op.ShowDialog() == DialogResult.OK)
            {
                xlWorkBook.SaveAs(op.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.DisplayAlerts = true;
                xlApp.Quit();

                MessageBox.Show("Votre fichier est enregistré sur le chemin suivant : c:\\Recrutement et suivi du personnel" + "test" + ".xls", "Informations", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
            {
                xlApp.Quit();
            }
                   
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

        private void ovalShape3_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void pictureBox3_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void pictureBox4_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void ovalShape7_Click(object sender, EventArgs e)
        {

        }
    }
}
