using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
//using static Tufl_CSV_To_ZZ.Decyzja;
namespace Tufl_CSV_To_ZZ
{
    public partial class WyborTrybuImportu : Form
    {
        static Button_Handle class1 = new Button_Handle();
        string connection_string = class1.GetConnectionString();
        
        
        public WyborTrybuImportu()
        {
            InitializeComponent();
            Load_Data();

        }
        public void Load_Data()
        {

            var select = $@"select r.Knt_Akronim 

                        from cdn.AtrybutyKlasy as k 

                               join cdn.Atrybuty as a on k.AtK_ID = Atr_AtkId 

                               join cdn.KntKarty as r on a.Atr_ObiNumer = r.Knt_GIDNumer and a.Atr_ObiTyp = Knt_GIDTyp 

                        where k.AtK_Nazwa = 'FormatImportuPlikuZZ'";
            using (SqlConnection connection = new SqlConnection(connection_string))
            {
                try
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(select, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    dataGridView1.ReadOnly = true;
                    dataGridView1.DataSource = dataSet.Tables[0];
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message.ToString(), "błąd zaciągania danych");
                }
                finally
                {
                    if (dataGridView1.Rows[0].Cells[0].Value != null)
                    {
                        textBox2.Text=dataGridView1.Rows[0].Cells[0].Value.ToString();
                    }
                    connection.Close();
                }
            }
            
            checkBox1.Checked = false;
            comboBox1.SelectedIndex = 0;
            label6.Text = Thread.CurrentThread.CurrentCulture.Name;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.Length > 0)
            {

                
                
                    
                    if (comboBox1.SelectedIndex != 0)
                    {
                    if (tB1.Value.ToShortDateString() == "")
                    {
                        MessageBox.Show("Proszę wypełnić numer dostawy");
                    }
                    else
                    {
                        string queryDecyzja = "if exists(SELECT TOP (1) 1 FROM [CDN].[Atrybuty] join cdn.AtrybutyKlasy on Atr_AtkId=Atk_id WHERE atk_Nazwa='Nr dostawy' and Atr_Wartosc='" + tB1.Value.ToShortDateString() + "') SELECT 'TAK' else SELECT 'NIE'";
                        string decyzja = class1.ConnectToSQL(connection_string, queryDecyzja);
                        if ((decyzja == "TAK" && comboBox1.SelectedIndex==1) || (decyzja == "NIE" && comboBox1.SelectedIndex == 2))
                        {
                            Decyzja dec_app = new Decyzja(dateTimePicker1.Value, textBox2.Text, comboBox1.SelectedIndex, class1, tB1.Value.ToShortDateString() == "" ? "NULL" : tB1.Value.ToShortDateString());
                            dec_app.Show();
                            //decyzja3 = decyzja2;
                            // dec_app.Dispose();
                            // dec_app.Close();

                        }
                        else
                        {
                            //MessageBox.Show("Jestem tutaj");
                            class1.ReadVariables(dateTimePicker1.Value, textBox2.Text == "" ? "NULL" : textBox2.Text, comboBox1.SelectedIndex, tB1.Value.ToShortDateString() == "" ? "NULL" : tB1.Value.ToShortDateString());
                            this.Close();
                        }
                    }
                    }
                    
                    else
                    {
                        if (textBox2.Text == "") textBox2.Text = "-";
                        switch (comboBox1.SelectedIndex)
                        {
                            case 0:

                                class1.ReadVariables(dateTimePicker1.Value, textBox2.Text == "" ? "NULL" : textBox2.Text, comboBox1.SelectedIndex, tB1.Value.ToShortDateString() == "" ? "NULL" : tB1.Value.ToShortDateString());
                                //MessageBox.Show($@"GID dostawcy {textBox2.Text} to {knt_gid[0].ToString()} a data to {dateTimePicker1.Value.ToString("yyMMdd")}");
                                //Application.OpenForms["PrepImport"].Close();
                                break;
                            case 1:
                                class1.ReadVariables(dateTimePicker1.Value, textBox2.Text == "" ? "NULL" : textBox2.Text, comboBox1.SelectedIndex, tB1.Value.ToShortDateString() == "" ? "NULL" : tB1.Value.ToShortDateString());
                                break;
                            case 2:
                                class1.ReadVariables(dateTimePicker1.Value, textBox2.Text == "" ? "NULL" : textBox2.Text, comboBox1.SelectedIndex, tB1.Value.ToShortDateString() == "" ? "NULL" : tB1.Value.ToShortDateString());
                                break;
                            default:
                                MessageBox.Show("Proszę wybrać tryb importu plików.");
                                break;
                        }


                        this.Close();
                    }
                    

                

            }
            else
            {
                MessageBox.Show("błąd! Nie wybrano żadnego dostawcy");
            }


        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
            {
                textBox2.Text = (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
            }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                textBox2.Text = "-";
            }
        }

        private void WyborTrybuImportu_Load(object sender, EventArgs e)
        {
            Button_Handle handle = new Button_Handle();
            if ( handle.getDatabase() == "ERPXL_Tuflowers_20231104")
            {
                button1.Visible = true;
                button1.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TestKlas test = new TestKlas();
            test.ShowDialog();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
