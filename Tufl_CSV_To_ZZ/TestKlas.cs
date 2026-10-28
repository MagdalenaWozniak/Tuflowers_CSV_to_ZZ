using cdn_api;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Tufl_CSV_To_ZZ.Button_Handle;

namespace Tufl_CSV_To_ZZ
{
    public partial class TestKlas : Form
    {
        public static int row_amount = 0;
        Button_Handle handle = new Button_Handle();
        private int return_value;
        private int XLtowar_id;
        private int sesja;
        private int XLWersjaAPI = 20231;
        private string path;
        private string connection_string;

        public TestKlas()
        {
            InitializeComponent();
        }

        private void TestKlas_Load(object sender, EventArgs e)
        {
            int loginkod = handle.XLLogin();
            richTextBox_Log_write(Color.Black, Color.White, $"Login do comarch, kod {loginkod}");

            FillDataGrid();
        }
        public void FillDataGrid()
        {
            //MessageBox.Show(ConnectToSQL(connection_string, "select count(*) from cdn.ONE_ZamMacierz where GIDTyp = " + gidtyp + " and GIDNumer = " + gidnumer + " and GIDlp = " + gidlp));
            this.dataGridView1.Rows.Clear();
            row_amount = row_amount = Int32.Parse(handle.ConnectToSQL(handle.GetConnectionString(), "select count(*) from dbo.FlowerGroup"));
            for (int i = 0; i < row_amount; i++)
            {
                var index = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[index].Cells["Nazwa"].Value = String.Format(handle.ConnectToSQL(handle.GetConnectionString(), "select NazwaKwiatu from dbo.FlowerGroup where Id="+i+1));
                this.dataGridView1.Rows[index].Cells["Grupa"].Value = String.Format(handle.ConnectToSQL(handle.GetConnectionString(), "select Grupa from dbo.FlowerGroup where Id=" + i + 1));
            }
        }
        private int[] getGrupa()
        {
            int[] rtrn_Group = new int[4] { 0, 0, 0, 0 };
            var resp = handle.ConnectToSQL(handle.GetConnectionString(), $@"select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Kod='{textBox4.Text}'");
            richTextBox_Log_write(ColorTranslator.FromHtml("#03A062"), Color.Black, $"Odnalezione GID to {resp}");

            string[] completeGrupa = resp.Split(';');
            if (completeGrupa.Length > 0 && textBox4.Text != "")
            {
                rtrn_Group[0] = Int32.Parse(completeGrupa[0]);
                rtrn_Group[1] = Int32.Parse(completeGrupa[1]);
                rtrn_Group[2] = Int32.Parse(completeGrupa[2]);
                rtrn_Group[3] = Int32.Parse(completeGrupa[3]);
            }
            else richTextBox_Log_write(Color.White, Color.DarkRed, $@"Błąd Nie znaleziono czegos dla towaru {textBox4.Text}, odpowiedz to {resp}");
            return rtrn_Group;
        }
        private int[] getGrupa(string grupa)
        {
            int[] rtrn_Group = new int[4] { 0, 0, 0, 0 };
            var resp = handle.ConnectToSQL(handle.GetConnectionString(), $@"select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Kod='{grupa}'");
            richTextBox_Log_write(ColorTranslator.FromHtml("#03A062"), Color.Black, $"Odnalezione GID to {resp}");

            string[] completeGrupa = resp.Split(';');
            if (completeGrupa.Length > 0 && grupa != "")
            {
                rtrn_Group[0] = Int32.Parse(completeGrupa[0]);
                rtrn_Group[1] = Int32.Parse(completeGrupa[1]);
                rtrn_Group[2] = Int32.Parse(completeGrupa[2]);
                rtrn_Group[3] = Int32.Parse(completeGrupa[3]);
            }
            else richTextBox_Log_write(Color.White, Color.DarkRed, $@"Błąd Nie znaleziono czegos dla towaru {textBox4.Text}, odpowiedz to {resp}");
            return rtrn_Group;
        }
        private void ReadCSV(string fileName)
        {

            List<CSVList> records = new List<CSVList>(); 
            try
            {
                StreamReader sr = new StreamReader(fileName);
                records = File.ReadAllLines(fileName)
                               .Skip(1)
                               .Select(v => CSVList.FromCsv(v))
                               .ToList();
                //CsvReader csv = new CsvReader(sr, System.Globalization.CultureInfo.GetCultureInfo("pl-PL"));
                //csv.Configuration.HasHeaderRecord = false;
                //records = csv.GetRecords<CSVList>().ToList();

            }
            catch (Exception e)
            {
                throw new Exception("Wyskoczył blad " + e);
            }
            var result = records
                         .Where(s => !string.IsNullOrWhiteSpace(s.Barcode))
                         .OrderBy(v => v.Barcode)
                         .GroupBy(x => x.Subdeb)
                         .OrderBy(y => y.Key)
                         .ToList();
 
                foreach (var Subdeb in result)
                {
                    if (Subdeb.FirstOrDefault().Invoicenumber.Length < 1 || Subdeb.FirstOrDefault().Barcode.Length < 1) continue; 
                         
                        for (int OrderCount = 0; OrderCount < Subdeb.Count(); OrderCount++)
                        {
                            AddNewIndex_XL(Subdeb.ElementAt(OrderCount), DateTime.Now, 1, "A");
                        }
                }
        }
        private void ReadNewCSV(string fileName)
        {

            List<newCSVList> records = new List<newCSVList>(); 
            try
            {
                StreamReader sr = new StreamReader(fileName);
                records = File.ReadAllLines(fileName)
                               .Skip(1)
                               .Select(v => newCSVList.FromCsv(v))
                               .ToList();
                //CsvReader csv = new CsvReader(sr, System.Globalization.CultureInfo.GetCultureInfo("pl-PL"));
                //csv.Configuration.HasHeaderRecord = false;
                //records = csv.GetRecords<CSVList>().ToList();

            }
            catch (Exception e)
            {
                throw new Exception("Wyskoczył blad " + e);
            }
            var result = records
                         .Where(s => !string.IsNullOrWhiteSpace(s.Barcode))
                         .OrderBy(v => v.Barcode)
                         .GroupBy(x => x.Subclient)
                         .OrderBy(y => y.Key)
                         .ToList();

            foreach (var Subdeb in result)
            {
                if (Subdeb.FirstOrDefault().Code.Length < 1 || Subdeb.FirstOrDefault().Barcode.Length < 1) continue;

                for (int OrderCount = 0; OrderCount < Subdeb.Count(); OrderCount++)
                {
                    AddNewIndex_XL(Subdeb.ElementAt(OrderCount), DateTime.Now, 1, "A");
                }
            }
        }
        private void TestKlas_FormClosed(object sender, FormClosedEventArgs e)
        {
          
        }

        private void button1_Click(object sender, EventArgs e)
        {

            OpenFileDialog oDlg = new OpenFileDialog();
            oDlg.Multiselect = true;
            errorCount = 0;
            if (DialogResult.OK == oDlg.ShowDialog())
            {
                path = oDlg.FileName;
            }
            
            connection_string = handle.GetConnectionString(); 
            pickedDate = DateTime.Now; 
            //richTextBox_Log_write(ColorTranslator.FromHtml("#03A062"), Color.Black, $@"Zalogowano do XL z kodem {xl_kod}"); 
            foreach (var FileName in oDlg.FileNames)
                {
                    try
                    {
                        
                        using (StreamReader sr = new StreamReader(FileName))
                        {
                            richTextBox_Log_write(Color.White, Color.Black, $@"--- Nowy plik {FileName} ---"); 
                            string fLine = sr.ReadLine();

                            string[] values = fLine.Split(';');

                            switch (values.Length)
                            {
                                case 23:
                                    ReadCSV(FileName);
                                    break;
                                case 18:

                                    ReadNewCSV(FileName);
                                    //richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,"Wylogowanie dało kod" + Logout());

                                    break;
                                default:
                                    richTextBox_Log_write(Color.White, Color.DarkRed, $@"Błąd! wybrano zły plik.");
                                    break;
                            }
                            richTextBox_Log_write(Color.White, Color.Black, $@"--- Zamkniętko plik {FileName} ---");
                            //XLLogin();
                            //XLogout();
                        }

                    }
                    catch (Exception ex)
                    {
                        //richTextBox_Log_write(Color.Black, Color.White, $@"Wylogowano! kod wylogowania {handle.XLLogout()}!");
                        throw new Exception($@"Wystąpił błąd przy zaciągu plików " + FileName + "\n" + ex);
                    }
                } 
        }
        public string Truncate(string value, int maxChars)
        {
            return value.Length <= maxChars ? value : value.Substring(0, maxChars);
        }
        string FindGroup(string FullAtricle)
        {
            string query = "select top 1 Grupa from dbo.FlowerGroup where NazwaKwiatu = ";
            string word = "";
            try
            {
                string[] ArticleArr = FullAtricle.Split(' ');
                for (int i = 0; i < ArticleArr.Count(); i++)
                {
                    for (int j = 0; j < ArticleArr.Count() - i; j++)
                    {
                        for (int k = 0; k <= i; k++)
                        {
                            if (k == 0)
                            {
                                query += "'" + ArticleArr[k + j];
                            }
                            else query += " " + ArticleArr[k + j];
                            if (k == i) query += "'";
                        }
                        if (i != ArticleArr.Count() - 1 && j != (ArticleArr.Count() - i)) query += " or NazwaKwiatu=";
                    }
                }
                //LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, query + "\n" + word);
                word = handle.ConnectToSQL(handle.GetConnectionString(), query);
            }
            catch (Exception e)
            {
                richTextBox1.Text += (Color.White, Color.DarkRed, $@"Błąd tworzenia zapytania o grupę towarową {e}");
                return word;
            }
            return word;
        }
        private int[] AddNewIndex_XL(newCSVList order, DateTime pickedTime, int tryb, string grupaPodatkowa)
        {
            //Stawka podatkowa z kodu Code - jest tabelka w XL 
            string nazwa = $@"{order.VBN}-{order.S1}-{order.S2}-{order.S3}_{pickedTime.ToString("yyMMdd")}_{"test"}";
            nazwa = Truncate(nazwa, 35);

            if (order.Subclient.Length == 0 || tryb == 2)
            {
                nazwa += "_TUFL"; 
            }
            else
                nazwa += ("_" + order.Subclient);
            XLTowarInfo_20231 xLTowarInfo = new XLTowarInfo_20231
            {
                Wersja = XLWersjaAPI,
                GIDFirma = 0,
                GIDLp = 0,
                GIDNumer = 0,
                GIDTyp = 0,
                Typ = 1,
                Kod = $"{nazwa}",
                //TwrGrupa = "KWIATY DEDYKOWANE", 
                StawkaPod = "23",
                GrupaPod = "A",
                StawkaPodSpr = "23",
                GrupaPodSpr = "A",
                Nazwa = order.Article,
                EAN = order.Barcode,
                CenaSpr = 100,
                PCN = order.Code.PadLeft(8, '0'),
                SWW = Regex.Replace(order.Code, ".{2}", "$0.").TrimEnd('.'),
                Jm = "szt.",
            };
            richTextBox_Log_write(Color.Black, Color.White, $@"Dodawanie nowego towaru o kodzie {nazwa} oraz nazwie {order.Article}");
            try
            {
                var dupa = getGrupa();
                xLTowarInfo.TwGNumer = dupa[0];
                xLTowarInfo.TwGTyp = dupa[1];
                xLTowarInfo.TwGLp = dupa[2];
                xLTowarInfo.TwGFirma = dupa[3];
                return_value = cdn_api.cdn_api.XLNowyTowar(handle.GetSession(), ref XLtowar_id, xLTowarInfo);
            }
            catch (Exception e)
            {
                throw new Exception("Wystapil wyjatek przy dodawaniu towaru. Kod API: " + return_value + ", Wyjątek: " + e);

            }
            if (return_value != 0)
            {
                richTextBox_Log_write(Color.White, Color.DarkRed, $"Błąd  dodania towaru {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");
                richTextBox_Log_write(Color.Black, Color.White, $"Poprwka wprowadzania grupy dla towaru {nazwa},"); 
                var dupa = getGrupa("KWIATY DEDYKOWANE");
                xLTowarInfo.TwGNumer = dupa[0];
                xLTowarInfo.TwGTyp = dupa[1];
                xLTowarInfo.TwGLp = dupa[2];
                xLTowarInfo.TwGFirma = dupa[3];
                return_value = cdn_api.cdn_api.XLNowyTowar(handle.GetSession(), ref XLtowar_id, xLTowarInfo);
                richTextBox_Log_write(Color.White, Color.DarkRed, $"Błąd poprawki towaru  {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} {'\n'} Przy grupie {dupa[0]} Kod API: {return_value}");

            }
            XLZamkniecieTowaruInfo_20231 xLZamkniecieTowaruInfo_20231 = new XLZamkniecieTowaruInfo_20231
            {
                Wersja = XLWersjaAPI,
                Tryb = return_value != 0 ? -1 : 0,
            };
            return_value = cdn_api.cdn_api.XLZamknijTowar(XLtowar_id, xLZamkniecieTowaruInfo_20231);
            if (return_value != 0)
            {
                richTextBox_Log_write(Color.White, Color.DarkRed, $"Błąd  zamknięcia towaru: {xLTowarInfo.Kod} dodany do magazynu o gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");
            } 
            return new int[4] { xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp, xLTowarInfo.GIDLp, xLTowarInfo.GIDFirma };
        }
        private int[] AddNewIndex_XL(CSVList order, DateTime pickedTime, int tryb, string grupaPodatkowa)
        {
            //Stawka podatkowa z kodu Code - jest tabelka w XL 
            string nazwa = $@"{order.Code}-{order.Mark1}-{order.Mark1}-{order.Mark1}_{pickedTime.ToString("yyMMdd")}_{"test"}";
            nazwa = Truncate(nazwa, 35);

            if (order.Subdeb.Length == 0 || tryb == 2)
            {
                nazwa += "_TUFL";
            }
            else
                nazwa += ("_" + order.Subdeb);
            XLTowarInfo_20231 xLTowarInfo = new XLTowarInfo_20231
            {
                Wersja = XLWersjaAPI,
                GIDFirma = 0,
                GIDLp = 0,
                GIDNumer = 0,
                GIDTyp = 0,
                Typ = 1,
                Kod = $"{nazwa}",
                //TwrGrupa = "KWIATY DEDYKOWANE", 
                StawkaPod = "23",
                GrupaPod = "A",
                StawkaPodSpr = "23",
                GrupaPodSpr = "A",
                Nazwa = order.Article,
                EAN = order.Barcode,
                CenaSpr = 100,
                PCN = order.Code.PadLeft(8, '0'),
                SWW = Regex.Replace(order.Code, ".{2}", "$0.").TrimEnd('.'),
                Jm = "szt.",
            };
            richTextBox_Log_write(Color.Black, Color.White, $@"Dodawanie nowego towaru o kodzie {nazwa} oraz nazwie {order.Article}");
           
            


            try
            {
                var dupa = getGrupa();
                xLTowarInfo.TwGNumer = dupa[0];
                xLTowarInfo.TwGTyp = dupa[1];
                xLTowarInfo.TwGLp = dupa[2];
                xLTowarInfo.TwGFirma = dupa[3];
                return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
            }
            catch (Exception e)
            {
                throw new Exception("Wystapil wyjatek przy dodawaniu towaru. Kod API: " + return_value + ", Wyjątek: " + e);

            }
            if (return_value != 0)
            {
                richTextBox_Log_write(Color.White, Color.DarkRed, $"Błąd  dodania towaru {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");
                richTextBox_Log_write(Color.Black, Color.White, $"Poprwka wprowadzania grupy dla towaru {nazwa} z gr. Kwiaty Dedykowane");
                var dupa = getGrupa("KWIATY DEDYKOWANE");
                xLTowarInfo.TwGNumer = dupa[0];
                xLTowarInfo.TwGTyp = dupa[1];
                xLTowarInfo.TwGLp = dupa[2];
                xLTowarInfo.TwGFirma = dupa[3];
                return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
                richTextBox_Log_write(Color.White, Color.DarkRed, $"Błąd poprawki towaru  {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} {'\n'} Przy grupie {dupa[0]} Kod API: {return_value}");

            }
            XLZamkniecieTowaruInfo_20231 xLZamkniecieTowaruInfo_20231 = new XLZamkniecieTowaruInfo_20231
            {
                Wersja = XLWersjaAPI,
                Tryb = return_value != 0 ? -1 : 0,
            };
            return_value = cdn_api.cdn_api.XLZamknijTowar(XLtowar_id, xLZamkniecieTowaruInfo_20231);
            if (return_value != 0)
            {
                richTextBox_Log_write(Color.White, Color.DarkRed, $"Błąd  zamknięcia towaru: {xLTowarInfo.Kod} dodany do magazynu o gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");
            }
            return new int[4] { xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp, xLTowarInfo.GIDLp, xLTowarInfo.GIDFirma };
        }
        public void richTextBox_Log_write(Color color1, Color color2, string text)
        {
            try
            {
                    int start = richTextBox1.Text.Length;
                    string rtbText = $"\n{(richTextBox1.Lines.Length + 1)}. " + text;
                    richTextBox1.AppendText(rtbText);
                    richTextBox1.Select(start, rtbText.Length);
                    if (color1 == null)
                    {
                        richTextBox1.SelectionColor = Color.Black;
                    }
                    else richTextBox1.SelectionColor = color1;
                    if (color2 != null)
                    {
                        richTextBox1.SelectionBackColor = color2;
                    }
            }
            catch (Exception e)
            {
                richTextBox1.AppendText("\n" + $"Błąd zaciągania tekstu do linii {richTextBox1.Lines.Length + 1}, tekst {text}, blad {e}");
            }
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
 
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
            {
                labelNazwa.Text = dataGridView1.Rows[e.RowIndex].Cells["Nazwa"].Value.ToString();
                labelGrupa.Text = dataGridView1.Rows[e.RowIndex].Cells["Grupa"].Value.ToString();
            }
        }
    }

}
