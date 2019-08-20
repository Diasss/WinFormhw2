using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                OleDbConnectionStringBuilder conn = new OleDbConnectionStringBuilder();
                conn.DataSource = textBox3.Text;
                conn.Provider = @"Provider = Microsoft.Jet.OLEDB.4.0";
                using (OleDbConnection olc = new OleDbConnection(conn.ConnectionString))
                {
                    try
                    {
                        olc.Open();
                        MessageBox.Show("Подключено к Access");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        olc.Close();
                    }
                }
            }
            if (checkBox2.Checked == true)
            {
                var conn = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source="+ textBox3.Text+"; Extended Properties=""Excel 8.0";

                using (OleDbConnection olc = new OleDbConnection(conn))
                {
                    try
                    {
                        olc.Open();
                        MessageBox.Show("Подключено к Excel");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        olc.Close();
                    }
                }
            }
            if (checkBox3.Checked == true)
            {
                SqlConnectionStringBuilder conn = new SqlConnectionStringBuilder();
                conn.DataSource = @"L206-3";
                conn.InitialCatalog = "ShopDB";
                conn.UserID = textBox1.Text;
                conn.Password = textBox2.Text;
                using (SqlConnection olc = new SqlConnection(conn.ConnectionString))
                {
                    try
                    {
                        olc.Open();
                        MessageBox.Show("Подключено к SQL");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        olc.Close();
                    }
                }
            }
        }
    }
}
