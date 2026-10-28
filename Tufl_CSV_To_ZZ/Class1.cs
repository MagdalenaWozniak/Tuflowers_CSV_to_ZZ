using System;
using System.Collections.Generic;
using System.Windows;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using cdn_api;
using Hydra;
using System.Windows.Forms;
using System.Data.SqlClient;
using ADODB;
using System.Data;
using System.IO;
using System.Security;
using System.Net;
using Microsoft.VisualBasic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
[assembly: CallbackAssemblyDescription(Name: "ONE_S_202_CSV_to_ZZ",
                                           Description: "Tworzenie ZZ oraz ZS z CSV",
                                           Author: "OneSupport",
                                           Version: "2.4",
                                           SystemVersion: "2021.0.2",
                                           Created: "02-04-2022")]

namespace Tufl_CSV_To_ZZ
{
    [SubscribeProcedure(Procedures.ZaNLista, "Tworzenie ZZ i ZS z CSV")]

    public class Button_Handle : Callback
    {
        public class FormData
        {
            public int[] zz_knt { get; set; }
            public DateTime orderTime { get; set; }
            public string dost_name { get; set; }

            
        }
        public class newCSVList
        {
            public string TotalAmnt { get; set; }
            public string VBN { get; set; }
            public string Article { get; set; }
            public string Subclient { get; set; }
            public string Box { get; set; }
            public string S1 { get; set; }
            public string S2 { get; set; }
            public string S3 { get; set; }
            public string RUPHEID { get; set; }
            public string InBos { get; set; }
            public string Grower { get; set; }
            public string Price { get; set; }
            public string TotalPrice { get; set; }
            public string ClientPrice { get; set; }
            public string ClientTotal { get; set; }
            public string Barcode { get; set; }
            public string Code { get; set; }
            public string ImgUrl { get; set; }
            public static newCSVList FromCsv(string csvLine)
            {
                string[] values = csvLine.Split(';');
                newCSVList csvlist = new newCSVList();
                csvlist.TotalAmnt = values[0];
                csvlist.VBN = values[1];
                csvlist.Article = values[2];
                csvlist.Subclient = values[3];
                csvlist.Box = values[4];
                csvlist.S1 = values[5];
                csvlist.S2 = values[6];
                csvlist.S3 = values[7];
                csvlist.RUPHEID = values[8];
                csvlist.InBos = values[9];
                csvlist.Grower = values[10];
                csvlist.Price = values[11];
                csvlist.TotalPrice = values[12];
                csvlist.ClientPrice = values[13];
                csvlist.ClientTotal = values[14];
                csvlist.Barcode = values[15];
                csvlist.Code = values[16];
                csvlist.ImgUrl = values[17];
                return csvlist;
            }
        }
        public class CSVList
        {
            public string Invoicenumber { get; set; }
            public string Date { get; set; }
            public string Clientnumber { get; set; }
            public string Group { get; set; }
            public string Article { get; set; }
            public string Quantity { get; set; }
            public string Quantitypercolli { get; set; }
            public string Price { get; set; }
            public string Mark1 { get; set; }
            public string Mark2 { get; set; }
            public string Mark3 { get; set; }
            public string Amount { get; set; }
            public string Origin { get; set; }
            public string Articlegroup { get; set; }
            public string Articleno { get; set; }
            public string Emballage { get; set; }
            public string Grower { get; set; }
            public string Subdeb { get; set; }
            public string Barcode { get; set; }
            public string Clientprice { get; set; }
            public string ClientVal { get; set; }
            public string Code { get; set; }
            public string Picture { get; set; }
            public static CSVList FromCsv(string csvLine)
            {
                string[] values = csvLine.Split(';');
                CSVList csvlist = new CSVList();
                csvlist.Invoicenumber = values[0];
                csvlist.Date = values[1];
                csvlist.Clientnumber = values[2];
                csvlist.Group = values[3];
                csvlist.Article = values[4];
                csvlist.Quantity = values[5];
                csvlist.Quantitypercolli = values[6];
                csvlist.Price = values[7];
                csvlist.Mark1 = values[8];
                csvlist.Mark2 = values[9];
                csvlist.Mark3 = values[10];
                csvlist.Amount = values[11];
                csvlist.Origin = values[12];
                csvlist.Articlegroup = values[13];
                csvlist.Articleno = values[14];
                csvlist.Emballage = values[15];
                csvlist.Grower = values[16];
                csvlist.Subdeb = values[17];
                csvlist.Barcode = values[18];
                csvlist.Clientprice = values[19];
                csvlist.ClientVal = values[20];
                csvlist.Code = values[21];
                csvlist.Picture = values[22];
                return csvlist;
            }
        }
        [DllImport("ClaRUN.dll")]
        public static extern void AttachThreadToClarion(bool x);

        ClaWindow elemButton;
        ClaWindow elemExport;

        OpenFileDialog oDlg = new OpenFileDialog();

        private string[] KlasyAtrybutuNazwa = new string[8] { "VBN", "Hodowca2", "Klient", "S1", "S2","S3","S4", "IMG_URL"};
        private string[] KlasyAtrybutuEsklepCiete = new string[4] { "Wysokość", "Gramatura", "Jakość", "Nakwitnięcie/Średnica kwiata" };
        private string[] KlasyAtrybutuEsklepDoniczkowe = new string[4] { "Wysokość", "Średnica", "Jakość", "Ilość kwiatów/pędów" };
        static List<Tuple<Color, Color, string>> infoLog = new List<Tuple<Color, Color, string>>();
        TextWriter tw2; 
        
        public static int errorCount = 0;
        public static int[] ZZ_knt = new int[4] { 0, 0, 0, 0 };
        public static DateTime pickedDate; 
        public static string delivery_name;
        public static string dostawa;
        public static string wlasnyAkronim = "DW";
        public static string clarionDate = "00000";
        public static int RodzajDzialania = 0;

        private int XLWersjaAPI = 20231;
        private int XLTrybWsadowy = 1;
        private int XLWinieta = -1;
        private int XLUtworzWlasnaSesje = 1;
        private static int XLZS_id = 1;
        private static int AltZS_id = -1;

        //private string XLSQLServer;
        //private string XLSQLDatabase;
        private static int sesja;
        private string query;

        private string connection_string;
        //private string XLProgramID = "ONE_S_20221_CSV_to_ZZ";
        private string path = @"C:\OneSupport\XML"; 
        //private string XLDatabase = "ERPXL_Tuflowers_20220908";
        private string XLDatabase = "ERPXL_Tuflowers";
        private string XLUser = "AUTOMAT";//"ADMIN";
        private string XLPass = "A$YLpkb9n.5uy";//"NE!w0rk12!";
        private string XLKey = "192.168.1.202::5000180665";//"SRV1::5000139268";//"SRV1::5000139268" //localhost::5000124544 albo "localhost::5000139271" ??
        private string XLConString = ";User ID=onesupport;Password=Ancr3d0$$db";

        /*
        private string XLDatabase = "Domar2";   // "Domar2";
        private string XLUser = "OS_AUTOM";     // "ADMIN";
        private string XLPass = "DomXL#53&";    // "Dupa2003!";
        private string XLConString = ";Integrated Security=SSPI;";//";User ID=SA;Password=DoMarSQL17";
        */

        private ClaWindow closeBtn; 

        HashSet<string> UniqueClients = new HashSet<string>();
        private int XLtowar_id;
        private int XLZam_id;

        public override void Init()
        { 
            AddSubscription(false, 0, Events.OpenWindow, new TakeEventDelegate(OnOpenwindow));
            AddSubscription(false, 0, Events.ResizeWindow, new TakeEventDelegate(Resize));
        }
        public string getDatabase()
        {
            return XLDatabase;
        }
        public int GetErrCount()
        {
            return errorCount;
        }
        public int GetSession()
        {
            return sesja;
        }
        public bool OnOpenwindow(Procedures ProcId, int ControlId, Events Event)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("pl-PL");
            Rectangle WndBnd = GetWindow().Bounds;
            elemButton = GetWindow().AllChildren["?TabZam"];
            GetWindow().OnBeforeSized += new TakeEventDelegate(Resize);
            elemExport = elemButton.Children.Add(ControlTypes.button);
            elemExport.Visible = true;
            elemExport.TextRaw = "Z CSV";
            elemExport.ToolTipRaw = "Wczytaj plik CSV na Z";
            closeBtn = GetWindow().AllChildren["?Close"];
            elemExport.Bounds = new Rectangle(closeBtn.Bounds.X, closeBtn.Bounds.Y - 25, 20, 15);
            //elemExport.OnAfterAccepted += new TakeEventDelegate(OpenForm);
            elemExport.OnAfterAccepted += new TakeEventDelegate(LoadCSV);
            elemExport.OnAfterAccepted += new TakeEventDelegate(OpenLog);
            //listID = GetWindow().AllChildren["?ListaZam"].Id;
            //popFiltr = GetWindow().AllChildren["?bUkryjFiltrDodatkowy"];
            //naduszenie przycisku
            //popFiltr.OnAfterAccepted += new TakeEventDelegate(Resize);
            return true;
        }
        public override void Cleanup()
        {
             
        }
        private bool Resize(Procedures ProcedureId, int ControlId, Events Event)
        {
            closeBtn = GetWindow().AllChildren["?Close"];
            elemExport.Bounds = new Rectangle(closeBtn.Bounds.X, closeBtn.Bounds.Y - 35, 20, 15);
            return true;
        }
        public void ReadVariables(  DateTime date, string textbox, int WybranyTryb, string nrDost)
        {
            try
            {
                int[] ZZ_knt_gid = ConnectToSQL(connection_string,
                    $@"if exists(select * from cdn.KntKarty where Knt_Akronim='{textbox}') 
                        select top 1 CONCAT(Knt_GIDNumer,';', Knt_GIDTyp,';', Knt_GIDLp,';', Knt_GIDFirma) from cdn.KntKarty where Knt_Akronim='{textbox}'
                        else select top 1 CONCAT(Knt_GIDNumer,';', Knt_GIDTyp,';', Knt_GIDLp,';', Knt_GIDFirma) from cdn.KntKarty where Knt_Akronim='ZS BRAK ATRYBUTU'").Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                   .Select(x => int.Parse(x))
                   .ToArray();
                ZZ_knt = ZZ_knt_gid;
                pickedDate = date;
                delivery_name = textbox;
                RodzajDzialania = WybranyTryb;
                dostawa = nrDost;
                
                switch (RodzajDzialania)
                {
                    case -1:
                        LogAppendValue(Color.Black, Color.White,$@"Tryb importu: Nieznany tryb, Wybrany dostawca {delivery_name} o GID {ZZ_knt[0]} z datą importu {pickedDate} ");
                        break;
                    case 0:
                        LogAppendValue(Color.Black, Color.White,$@"Wybrano tryb: Domyślne ZZ->ZS, Wybrany dostawca {delivery_name} o GID {ZZ_knt[0]} z datą importu {pickedDate} ");
                        break;
                    case 1:
                        LogAppendValue(Color.Black, Color.White,$@"Wybrano tryb: Tylko ZS, Wybrany dostawca {delivery_name} o GID {ZZ_knt[0]} z datą importu {pickedDate} ");
                        break;
                    case 2:
                        LogAppendValue(Color.Black, Color.White,$@"Wybrano tryb: Tylko ZZ, Wybrany dostawca {delivery_name} o GID {ZZ_knt[0]} z datą importu {pickedDate} ");
                        break;
                    default:
                        LogAppendValue(Color.Black, Color.White,$@"Tryb importu: Nieznany tryb, Wybrany dostawca {delivery_name} o GID {ZZ_knt[0]} z datą importu {pickedDate} ");
                        break;
                }
                
            }
            catch (Exception e)
            {
                LogAppendValue(Color.White, Color.DarkRed,$@"Błąd przy imporcie danych dostawcy i/lub daty: {ZZ_knt[0]} oraz {date.ToString("yyMMdd")} \nWyjątek: {e.Message}");
            }
 
        }
        public bool OpenLog(Procedures ProcId, int ControlId, Events Event)
        {
            ShowLog newLog = new ShowLog();
            newLog.richTextBox_Log_write(infoLog);
            newLog.Show();
            return true;
        }
        public void LoadNewCSV()
        {
             
            Cursor.Current = Cursors.WaitCursor;

            oDlg.Multiselect = true;
            errorCount = 0;
            if (DialogResult.OK == oDlg.ShowDialog())
            {
                path = oDlg.FileName;
            } 
            string xl_kod =XLLogin().ToString();
            connection_string = GetConnectionString();
            WyborTrybuImportu wyborTrybu = new WyborTrybuImportu();
            pickedDate = DateTime.Now;
            RodzajDzialania = -1;
            LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $@"Zalogowano do XL z kodem {xl_kod}");
            wyborTrybu.ShowDialog();
            if (RodzajDzialania != -1)
            {
                foreach (var FileName in oDlg.FileNames)
                {
                    try
                    {
                        LogAppendValue(Color.White, Color.Black,$@"--- Nowy plik {FileName} ---");
                        StreamReader sr = new StreamReader(FileName);
                        string fLine = sr.ReadLine();
                        //List<CSVList> records;
                        //List<newCSVList> newrecords;

                        string[] values = fLine.Split(';');

                        switch (values.Length)
                        {
                            case 23:
                                ReadCSV(FileName);
                                break;
                            case 18:

                                ReadnewCSV(FileName);
                                //richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,"Wylogowanie dało kod" + Logout());

                                break;
                            default:
                                LogAppendValue(Color.White, Color.DarkRed, $@"Błąd! wybrano zły plik.");
                                break;
                        }
                        LogAppendValue(Color.White, Color.Black,$@"--- Zamkniętko plik {FileName} ---");
                        //XLLogin();
                        //XLogout();
                    }
                    catch (Exception e)
                    {
                        LogAppendValue(Color.Black, Color.White,$@"Wylogowano! kod wylogowania {XLLogout()}!");
                        throw new Exception($@"Wystąpił błąd przy zaciągu plików " + FileName +"\n"+e);
                    } 
                }
            }
            int odp = XLLogout();
            if (odp != 0)
            {
                LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $@"Wylogowano z błędnym kodem wylogowania {odp}!");
                
            }
        }
        public bool LoadCSV(Procedures ProcId, int ControlId, Events Event)
        {
             
            infoLog.Clear();
            tw2= new StreamWriter(@"C:\log\log.txt", append: true);
            tw2.WriteLine();
            tw2.WriteLine();
            tw2.WriteLine(DateTime.Now.ToString());
            Cursor.Current = Cursors.WaitCursor;
            OpenFileDialog oDlg = new OpenFileDialog();
            oDlg.Multiselect = true;
            errorCount = 0;
            if (DialogResult.OK == oDlg.ShowDialog())
            {
                path = oDlg.FileName;
            }

            string xl_kod = XLLogin().ToString();
            connection_string = GetConnectionString();
            LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $@"Zalogowano do XL z kodem {xl_kod}");
            WyborTrybuImportu wyborTrybu = new WyborTrybuImportu();
            pickedDate = DateTime.Now;
            RodzajDzialania = -1;
           
            wyborTrybu.ShowDialog();
            if(RodzajDzialania!=-1)
            { 
                foreach (var FileName in oDlg.FileNames)
                {
                    
                    try
                    {
                        
                        LogAppendValue(Color.Black, Color.White, $@"--- Nowy plik {FileName} ---" + Environment.NewLine);
                        StreamReader sr = new StreamReader(FileName);
                        string fLine = sr.ReadLine();
                        //List<CSVList> records;
                        //List<newCSVList> newrecords;

                        string[] values = fLine.Split(';');
                       


                        switch (values.Length)
                        {
                            case 23:
                                {
                                    
                                    ReadCSV(FileName);
                                    break;
                                }
                            case 18:
                                ReadnewCSV(FileName);
                                //richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,"Wylogowanie dało kod" + Logout());

                                break;
                            default:
                                LogAppendValue(Color.White, Color.DarkRed,$@"Wystąpił błąd! wybrano zły plik.");
                                break;
                        }
                        LogAppendValue(Color.Black, Color.White,$@"--- Zamkniętko plik {FileName} ---"+Environment.NewLine);
                        sr.Close();
                        //XLLogin();
                        //XLogout();
                    }
                    catch (Exception e)
                    {
                        LogAppendValue(Color.Black, Color.White,$@"Wylogowano! kod wylogowania {XLLogout()}!");
                        throw new Exception("Wyskoczył błąd load CSV" + e);
                    }
                }
            } else
 
            if (XLLogout() != 0)
            {
            MessageBox.Show($@"Wylogowano z błędnym kodem wylogowania!");
            }
            tw2.Close();
            return true;
        }
        private List<CSVList> ReadCSV(string fileName)
        {

            
            

            List<CSVList> records = new List<CSVList>();
            double cnt = 0;
            try
            {
                StreamReader sr = new StreamReader(fileName);
                records = File.ReadAllLines(fileName)
                               .Skip(1)
                               .Where(s => !string.IsNullOrWhiteSpace(s))
                               .Select(v => CSVList.FromCsv(v))
                               .ToList();
                //CsvReader csv = new CsvReader(sr, System.Globalization.CultureInfo.GetCultureInfo("pl-PL"));
                //csv.Configuration.HasHeaderRecord = false;
                //records = csv.GetRecords<CSVList>().ToList();

            }
            catch (Exception e)
            {
                throw new Exception("Wyskoczył blad read CSV" + e);
            }

            //records.FindIndex(s => s.Barcode == "");

            //https://social.msdn.microsoft.com/Forums/en-US/1d36373e-3631-4c5b-96c4-cc80605ea149/add-new-column-in-existing-csv-file-using-c?forum=aspgettingstarted

            var result = records
                         .Where(s => !string.IsNullOrWhiteSpace(s.Barcode)) 
                         .OrderBy(v => v.Barcode)
                         .GroupBy(x => x.Subdeb)
                         .OrderBy(y => y.Key)
                         .ToList();
            Pasek pp = new Pasek();
            pp.Show();
            pp.pBar1.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            pp.pBar1.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            pp.pBar1.Maximum = result.Count() + 1;
            // Set the initial value of the ProgressBar.
            pp.pBar1.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            pp.pBar1.Step = 1;


            if (RodzajDzialania != 1)
            {
                foreach (var Subdeb in result)
                {
                    pp.pBar1.PerformStep();
                    pp.label2.Text = Subdeb.Key;
                    try
                    {
                        if (Subdeb.FirstOrDefault().Invoicenumber.Length < 1 || Subdeb.FirstOrDefault().Barcode.Length < 1) continue;

                       
                        if (Subdeb.Key == "")
                        {
                            query = $@"if exists(select top 1 ZaN_GIDNumer from cdn.ZamNag where ZaN_DokumentObcy = {Subdeb.First().Invoicenumber} )
                        select top 1 1
                        else 
                        select top 1 0 ";
                            string msg = ConnectToSQL(connection_string, query);
                            if (msg == "1")
                            {
                                LogAppendValue(Color.Black, Color.White, $@"Istnieje już taka faktura z numerem {Subdeb.First().Invoicenumber}");
                                //continue;
                            }


                            int _barcodeCount = 0;
                            int[] ZZ_GId = new int[4] { 0, 0, 0, 0 };
                            int[] gid_knt = new int[4] { 0, 0, 0, 0 };
                            string grupaPodatkowa = "A";
                            for (int OrderCount = 0; OrderCount < Subdeb.Count(); OrderCount++)
                            {

                               
                                _barcodeCount = 0;
                                cnt = 0;
                                if (OrderCount == 0)
                                {
                                    try
                                    {
                                        gid_knt = ConnectToSQL(connection_string, $"select top 1 CONCAT(Knt_GIDNumer,';', Knt_GIDTyp,';', Knt_GIDLp,';', Knt_GIDFirma) from cdn.KntKarty where Knt_Akronim='{(Subdeb.ElementAt(OrderCount).Subdeb == "" ? Subdeb.ElementAt(OrderCount).Clientnumber : Subdeb.ElementAt(OrderCount).Subdeb)}'")
                                            .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                            .Select(x => int.Parse(x))
                                            .ToArray();

                                    }
                                    catch (Exception e)
                                    {
                                        LogAppendValue(Color.White, Color.DarkRed, "Wystąpił błąd przy pobieraniu kontrahenta przy zamowieniu wlasnym " + gid_knt[0] + " komunikat " + e);
                                        continue;
                                    }
                                    finally
                                    {
                                        grupaPodatkowa = ConnectToSQL(connection_string, $@"select top 1 Knt_ExpoKraj from cdn.KntKarty where Knt_GIDNumer={ZZ_knt[0]}");
                                        ZZ_GId = CreateZZ(ZZ_knt, Subdeb.ElementAt(OrderCount).Grower, Subdeb.ElementAt(OrderCount).Invoicenumber, false, DateTime.Parse(Subdeb.ElementAt(OrderCount).Date), pickedDate, grupaPodatkowa);
                                    }


                                }
                                try
                                {
                                    if (Double.TryParse(Subdeb.ElementAt(OrderCount).Quantity, out double temp) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                    cnt += temp;
                                    try
                                    {
                                        if (OrderCount < (Subdeb.Count() - 1))
                                        {
                                            while (Subdeb.ElementAt(OrderCount).Barcode == Subdeb.ElementAt(OrderCount + 1).Barcode &&
                                                OrderCount < (Subdeb.Count() - 1) &&
                                                Subdeb.ElementAt(OrderCount).Price == Subdeb.ElementAt(OrderCount + 1).Price &&
                                                Subdeb.ElementAt(OrderCount).Article == Subdeb.ElementAt(OrderCount + 1).Article)
                                            {
                                                _barcodeCount++;
                                                OrderCount++;
                                                if (Double.TryParse(Subdeb.ElementAt(OrderCount).Quantity, out double temp2) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                                cnt += temp2;
                                                if (OrderCount == (Subdeb.Count() - 1))
                                                {
                                                    LogAppendValue(Color.Black, Color.White, $@"Order count: {OrderCount} a Subdeb count {Subdeb.Count()} więc robie break (ostatnia wartość duplikatu)");
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex2)
                                    {
                                        LogAppendValue(Color.White, Color.DarkRed, $@"BŁAD {++errorCount} przy dodawaniu duplikatow dla  order count {OrderCount} przy {Subdeb.ElementAt(OrderCount).Article}  {ex2}");
                                    }
                                    finally
                                    {
                                        if (cnt > 0)
                                        {
                                            if (_barcodeCount > 0) LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} z Barcode {Subdeb.ElementAt(OrderCount).Barcode} powtórzył się {_barcodeCount + 1} razy");
                                            else LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} z Barcode {Subdeb.ElementAt(OrderCount).Barcode}");
                                            
                                            if(RodzajDzialania==2)
                                                AddElement_XL_ZZ(XLZam_id, Subdeb.ElementAt(OrderCount), cnt.ToString(), grupaPodatkowa, "ZZ");
                                            else
                                                AddElement_XL(XLZam_id, Subdeb.ElementAt(OrderCount), AddIndex_XL(Subdeb.ElementAt(OrderCount), pickedDate, RodzajDzialania, grupaPodatkowa), cnt.ToString(), grupaPodatkowa, "ZZ");
                                        
                                        
                                        }
                                        cnt = 0;
                                    }

                                }
                                /*
                                try
                                {
                                    for (int i = 0; i < UniquePositions.Count; i++)
                                    {
                                        //(UniquePositions[i].Amount);
                                        if (Double.TryParse(UniquePositions[i].Quantity, out double temp) == false) LogAppendValue(Color.White, Color.DarkRed,$@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                        cnt += temp;
                                        if (i > 0) OrderCount++;
                                    }
                                    if (cnt > 0)
                                    {
                                        //Zakładanie indeksu
                                        AddElement_XL(XLZam_id, Subdeb.ElementAt(OrderCount), AddIndex_XL(Subdeb.ElementAt(OrderCount), pickedDate, RodzajDzialania, grupaPodatkowa), cnt.ToString(), grupaPodatkowa);
                                   
                                    }

                                }
                                */
                                catch (Exception e)
                                {

                                    throw new Exception("blad dodawania indeksów i zakladania pozycji: " + e);
                                }
                            }
                            CloseZZ(ZZ_knt, "duuvenworde", ZZ_GId, XLZam_id, DateTime.Parse(Subdeb.FirstOrDefault().Date), pickedDate);
                        }
                        else if (Subdeb.Key != "")
                        {
                            query = $@"if not exists(select top 1 Atr_AtrNumer  from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId=AtK_ID where Atr_Wartosc='{Subdeb.Key}' and AtK_Nazwa LIKE 'KOD %') 
                        select top 1 2 
                        else 
                        if exists(select top 1 ZaN_GIDNumer from cdn.ZamNag where ZaN_DokumentObcy = '{Subdeb.First()}')
                        select top 1 1
                        else 
                        select top 1 0 ";
                            string msg = ConnectToSQL(connection_string, query);
                            //MessageBox.Show("Msg: " + msg, "Komunikat");
                            if (msg == "2")
                            {
                                LogAppendValue(Color.White, Color.DarkRed, $"Błąd! brak klienta w bazie z atrybutem dla kodu {Subdeb.Key}");
                                continue;
                            }
                            else if (msg == "1")
                            {
                                LogAppendValue(Color.Black, Color.White, "Istnieje już taka faktura.");
                                continue;
                            }
                            int[] ZZ_gid = new int[4] { 0, 0, 0, 0 };
                            int _barcodeCount = 0;
                            string grupaPod = "A";
                            for (int OrderCount = 0; OrderCount < Subdeb.Count(); OrderCount++)
                            {
                                cnt = 0;
                                _barcodeCount = 0;
                                if (OrderCount == 0)
                                {
                                    try {
                                        grupaPod = ConnectToSQL(connection_string, $@"select top 1 Knt_ExpoKraj from cdn.KntKarty where Knt_GIDNumer={ZZ_knt[0]}");
                                        ZZ_gid = CreateZZ(ZZ_knt, Subdeb.ElementAt(OrderCount).Grower, Subdeb.ElementAt(OrderCount).Invoicenumber, true, DateTime.Parse(Subdeb.ElementAt(OrderCount).Date), pickedDate, grupaPod);

                                    }
                                    catch(Exception e)
                                    {
                                        MessageBox.Show("Błąd podczas dodawania ZZ (wywołanie funkcji): " + e.Message, "Komunikat");
                                    }
                                }
                                try
                                {
                                    if (Double.TryParse(Subdeb.ElementAt(OrderCount).Quantity, out double temp) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                    cnt += temp;
                                    if (OrderCount < Subdeb.Count() - 1)
                                    {
                                       

                                            while (Subdeb.ElementAt(OrderCount).Barcode == Subdeb.ElementAt(OrderCount + 1).Barcode &&
                                                    OrderCount < (Subdeb.Count() - 1) &&
                                                    Subdeb.ElementAt(OrderCount).Price == Subdeb.ElementAt(OrderCount + 1).Price &&
                                                    Subdeb.ElementAt(OrderCount).Article == Subdeb.ElementAt(OrderCount + 1).Article)
                                            {
                                                _barcodeCount++;
                                                OrderCount++;
                                                if (Double.TryParse(Subdeb.ElementAt(OrderCount).Quantity, out double temp2) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                                cnt += temp2;
                                                if (OrderCount == Subdeb.Count() - 1)
                                                    break;

                                                /* if (Subdeb.Key == "75CF")
                                                 {
                                                     MessageBox.Show("Quantity: " + Subdeb.ElementAt(OrderCount).Quantity + "\nOrder count: " + OrderCount, "Komunikat");
                                                     MessageBox.Show("Price: " + Subdeb.ElementAt(OrderCount + 1).Price + "\nOrder count: " + OrderCount, "Komunikat");
                                                     MessageBox.Show("Article: " + Subdeb.ElementAt(OrderCount + 1).Article + "\nOrder count: " + OrderCount, "Komunikat");
                                                     MessageBox.Show("Barcode: " + Subdeb.ElementAt(OrderCount + 1).Barcode + "\nOrder count: " + OrderCount, "Komunikat");
                                                 }*/
                                            }
                                        
                                    }
                                    
                                        
                                      

                                    if (cnt > 0)
                                    {
                                        if (_barcodeCount > 0) LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} z Barcode {Subdeb.ElementAt(OrderCount).Barcode} powtórzył się {_barcodeCount + 1} razy");
                                        else LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} z Barcode {Subdeb.ElementAt(OrderCount).Barcode}");
                                        //MessageBox.Show("Jestem tutaj :)", "Komunikat");
                                        if(RodzajDzialania==2)
                                            AddElement_XL_ZZ(XLZam_id, Subdeb.ElementAt(OrderCount),  cnt.ToString(), grupaPod, "ZZ");
                                        else
                                            AddElement_XL(XLZam_id, Subdeb.ElementAt(OrderCount), AddIndex_XL(Subdeb.ElementAt(OrderCount), pickedDate, RodzajDzialania, grupaPod), cnt.ToString(), grupaPod, "ZZ");

                                    }
                                    cnt = 0;
                                }
                                catch(Exception e)
                                {
                                    MessageBox.Show("Błąd dodawania elementów w ZZ: " + e.Message, "Komunikat");
                                }
                                /*
                                try
                                {
                                    for (int i = 0; i < UniquePositions.Count; i++)
                                    {
                                        //(UniquePositions[i].Quantity);
                                        if (Double.TryParse(UniquePositions[i].Quantity, out double temp) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                        cnt += temp;
                                        if (i > 0) OrderCount++;
                                    }
                                    if (cnt > 0)
                                    {
                                        //Zakładanie indeksu
                                        AddElement_XL(XLZam_id, Subdeb.ElementAt(OrderCount), AddIndex_XL(Subdeb.ElementAt(OrderCount), pickedDate, RodzajDzialania, grupaPod), cnt.ToString(), grupaPod);
                                    }

                                }
                                */
                                    /*catch (Exception en)
                                    {

                                        throw new Exception("blad: " + en);
                                    }*/
                                }
                            //XLZS_id 
                            CloseZZ(ZZ_knt, wlasnyAkronim, ZZ_gid, XLZam_id, DateTime.Parse(Subdeb.FirstOrDefault().Date), pickedDate);
                            if (RodzajDzialania != 2) CreateZS(ZZ_gid, Subdeb.Key, pickedDate, pickedDate);
                        }
                    }
                    catch(Exception e)
                    {
                        MessageBox.Show("Błąd w readCSV rodzaj dzialania!=1" + e.Message+"\n"+Subdeb.Key);
                    }
                }
            }
            else
            {
                //AlternatywnyZS
                foreach (var Subdeb in result)
                {
                    pp.pBar1.PerformStep();
                    pp.label2.Text = Subdeb.Key;
                    try
                    {
                        if (Subdeb.FirstOrDefault().Invoicenumber.Length < 1 || Subdeb.FirstOrDefault().Barcode.Length < 1) continue;
                        int rtrn_val = -1;
                        int _barcodeCount = 0;
                        int[] new_knt = new int[4] { 0, 0, 0, 0 };
                        cnt = 0;
                        int AltZS_id = -1;
                        XLDokumentZamNagInfo_20231 xlNewDokumentZamAltInfo = new XLDokumentZamNagInfo_20231();
                        for (int OrderCount = 0; OrderCount < Subdeb.Count(); OrderCount++)
                        {
                            _barcodeCount = 0;
                            cnt = 0;
                            if (OrderCount == 0)
                            {
                                string odp = "";
                                try
                                {
                                   
                                    //Zmienia się KNT ze względu na tryb
                                    //odp = ConnectToSQL(connection_string, $"select top 1 CONCAT(Atr_ObiNumer,';', Atr_ObiTyp,';', Atr_ObiLp,';', Atr_ObiFirma)  from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId=AtK_ID where Atr_Wartosc='{(Subdeb.ElementAt(OrderCount).Subdeb == "" ? wlasnyAkronim : Subdeb.ElementAt(OrderCount).Subdeb)}' and AtK_Nazwa LIKE 'KOD %'");
                                    try
                                    {
                                       
                                        new_knt = ConnectToSQL(connection_string, $"if exists(select * from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId=AtK_ID where Atr_Wartosc='{(Subdeb.ElementAt(OrderCount).Subdeb.Length < 1 ? wlasnyAkronim : Subdeb.ElementAt(OrderCount).Subdeb)}' and AtK_Nazwa LIKE 'KOD %')" +
                                            $" select top 1 CONCAT(Atr_ObiNumer,';', Atr_ObiTyp,';', Atr_ObiLp,';', Atr_ObiFirma)  from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId=AtK_ID where Atr_Wartosc='{(Subdeb.ElementAt(OrderCount).Subdeb.Length < 1 ? wlasnyAkronim : Subdeb.ElementAt(OrderCount).Subdeb)}' and AtK_Nazwa LIKE 'KOD %'" +
                                            $" else select top 1 CONCAT(Knt_GIDNumer,';',Knt_GIDTyp,';', Knt_GIDLp,';',Knt_GIDFirma)  from cdn.KntKarty  where KNt_Akronim='ZS BRAK ATRYBUTU'")
                                           //+ $" AtK_Nazwa = 'KOD ' + '(select top 1 Knt_Nazwa1 from cdn.KntKarty where Knt_GIDNumer = {knt[0]} and Knt_GIDTyp = {knt[1]} and Knt_GIDLp = {knt[2]} and Knt_GIDFirma = {knt[3]})'")
                                           .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                           .Select(x => int.Parse(x))
                                           .ToArray();
                                        if (new_knt[0] <= 0)
                                        {
                                            LogAppendValue(Color.White, Color.DarkRed, $@"nie odnaleziono kontrahenta {(Subdeb.ElementAt(OrderCount).Subdeb == "" ? wlasnyAkronim : Subdeb.ElementAt(OrderCount).Subdeb)} do ZS {new_knt[0]} !");

                                           
                                            break;
                                        }
                                        else LogAppendValue(Color.Black, Color.White, $@"dla KNT {(Subdeb.ElementAt(OrderCount).Subdeb == "" ? wlasnyAkronim : Subdeb.ElementAt(OrderCount).Subdeb)} odnaleziono gid {new_knt[0]}");
                                    }
                                    catch (Exception e)
                                    {
                                        LogAppendValue(Color.White, Color.DarkRed, $@"nie odnaleziono kontrahenta {(Subdeb.ElementAt(OrderCount).Subdeb == "" ? wlasnyAkronim : Subdeb.ElementAt(OrderCount).Subdeb)} do ZS (błąd split do array)  + {e}!");
                                        break;
                                    }
                                    xlNewDokumentZamAltInfo.Wersja = XLWersjaAPI;
                                    xlNewDokumentZamAltInfo.Typ = 6; //ZS
                                    xlNewDokumentZamAltInfo.Tryb = 2;
                                    xlNewDokumentZamAltInfo.KntFirma = new_knt[3];
                                    xlNewDokumentZamAltInfo.KntLp = new_knt[2];
                                    xlNewDokumentZamAltInfo.KntNumer = new_knt[0];
                                    xlNewDokumentZamAltInfo.KntTyp = new_knt[1];
                                    //xlNewDokumentZamAltInfo.Rok = pickedDate.Year;
                                    //xlNewDokumentZamAltInfo.Miesiac = pickedDate.Month;
                                    //xlNewDokumentZamAltInfo.DataWystawienia = GetClarionDate(pickedDate);         // ewentualnie GetClarionDate(pickedDate)
                                    //xlNewDokumentZamAltInfo.DataPotwierdzeniaOferty = GetClarionDate(pickedDate);
                                    xlNewDokumentZamAltInfo.DataPotwierdzeniaOferty = GetClarionDate(DateTime.Parse(Subdeb.ElementAt(OrderCount).Date));
                                    xlNewDokumentZamAltInfo.DataWystawienia = GetClarionDate(DateTime.Parse(Subdeb.ElementAt(OrderCount).Date));
                                    xlNewDokumentZamAltInfo.DataRealizacji= GetClarionDate(DateTime.Parse(Subdeb.ElementAt(OrderCount).Date));
                                    xlNewDokumentZamAltInfo.DataWaznosci= GetClarionDate(DateTime.Parse(Subdeb.ElementAt(OrderCount).Date));
                                    xlNewDokumentZamAltInfo.FlagaNB = "N";
                                    xlNewDokumentZamAltInfo.ZamSeria = Subdeb.ElementAt(OrderCount).Subdeb != "" ? "DYSTR" : "CENTR";
                                    xlNewDokumentZamAltInfo.Waluta = "PLN";

                                    rtrn_val = cdn_api.cdn_api.XLNowyDokumentZam(sesja, ref AltZS_id, xlNewDokumentZamAltInfo);

                                    


                                }
                                catch (Exception ex)
                                {

                                    LogAppendValue(Color.Black, Color.White, $@"Nie odnaleziono kontrahenta! błąd to {ex} a odp na wybór kontrahenta to {odp}");
                                    XLZamkniecieDokumentuZamInfo_20231 xLZamkniecieDokumentuInfo_202022 = new XLZamkniecieDokumentuZamInfo_20231
                                    {
                                        Wersja = XLWersjaAPI,
                                        TrybZamkniecia = 1,
                                    };
                                    break;
                                }

                                if (rtrn_val != 0 || new_knt[0] == 0)
                                {

                                    if (rtrn_val != 33)
                                    {
                                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount}. zakładania dokument ZS o gid {xlNewDokumentZamAltInfo.GIDNumer} z kodem {rtrn_val} i knt {xlNewDokumentZamAltInfo.KntNumer} oraz {odp}");
                                    }
                                    else
                                    {
                                        LogAppendValue(Color.White, Color.DarkRed,
                                            $@"Błąd API {rtrn_val} : Brak licencji na moduł Zamówienia - prawdopodobnie potrzebne ponowne zalogowanie do XL");
                                    }


                                    XLZamkniecieDokumentuZamInfo_20231 xLZamkniecieDokumentuInfo_202022 = new XLZamkniecieDokumentuZamInfo_20231
                                    {
                                        Wersja = XLWersjaAPI,
                                        TrybZamkniecia = 1,
                                    };
                                    break;
                                }
                                else LogAppendValue(Color.White, Color.DarkRed, $@"Utworzono ZS o GID {xlNewDokumentZamAltInfo.GIDNumer}");
                            }
                            try
                            {
                                if (Double.TryParse(Subdeb.ElementAt(OrderCount).Quantity, out double temp) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                cnt += temp;
                                if (OrderCount < Subdeb.Count() - 1)
                                {
                                    while (Subdeb.ElementAt(OrderCount).Barcode == Subdeb.ElementAt(OrderCount + 1).Barcode &&
                                                OrderCount < (Subdeb.Count() - 1) &&
                                                Subdeb.ElementAt(OrderCount).Price == Subdeb.ElementAt(OrderCount + 1).Price &&
                                                Subdeb.ElementAt(OrderCount).Article == Subdeb.ElementAt(OrderCount + 1).Article)
                                    {
                                        _barcodeCount++;
                                        OrderCount++;
                                        if (Double.TryParse(Subdeb.ElementAt(OrderCount).Quantity, out double temp2) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                        cnt += temp2;
                                        if (OrderCount == (Subdeb.Count() - 1))
                                        {
                                            LogAppendValue(Color.Red, Color.White, $@"Order count: {OrderCount} a Subdeb count {Subdeb.Count()}. Przerywam (ostatni duplikat)");
                                            break;
                                        }
                                    }
                                }
                                if (cnt > 0)
                                {
                                    if (_barcodeCount > 0) LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} z Barcode {Subdeb.ElementAt(OrderCount).Barcode} powtórzył się {_barcodeCount + 1} razy");
                                    else LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} z Barcode {Subdeb.ElementAt(OrderCount).Barcode}");
                                    AddElement_XL(AltZS_id, Subdeb.ElementAt(OrderCount), AddIndex_XL(Subdeb.ElementAt(OrderCount), pickedDate, RodzajDzialania, "A"), cnt.ToString(), "A", "ZS");

                                }
                                cnt = 0;
                            }

                            catch (Exception e)
                            {

                                throw new Exception("blad: " + e);
                            }
                        }
                        if (rtrn_val != -1)
                        {
                            XLZamkniecieDokumentuZamInfo_20231 xLZamkniecieDokumentuInfo_20202 = new XLZamkniecieDokumentuZamInfo_20231
                            {
                                Wersja = XLWersjaAPI,
                                TrybZamkniecia = rtrn_val == 0 ? 0 : 1,
                            };
                            rtrn_val = cdn_api.cdn_api.XLZamknijDokumentZam(AltZS_id, xLZamkniecieDokumentuInfo_20202);
                            if (xLZamkniecieDokumentuInfo_20202.TrybZamkniecia == 0)
                            {
                                if (DateTime.Parse(Subdeb.FirstOrDefault().Date) != pickedDate)
                                {
                                    var str = ChangeDate(GetClarionDate(DateTime.Parse(Subdeb.FirstOrDefault().Date)), pickedDate, xlNewDokumentZamAltInfo.GIDNumer);

                                    if (str != "")
                                    {
                                        LogAppendValue(Color.Black, Color.White, $@"Bład przy zmianie daty dostawy na {GetClarionDate(DateTime.Parse(Subdeb.FirstOrDefault().Date))}! {str}");
                                    }
                                    else LogAppendValue(Color.Black, Color.White, $@"Pomyślnie zmieniono date na {GetClarionDate(DateTime.Parse(Subdeb.FirstOrDefault().Date))}");


                                }
                                string zz_atr_new = $@"if not exists (select 1 from cdn.Atrybuty left join cdn.AtrybutyKlasy on AtK_ID=Atr_AtkId where Atr_ObiTyp = {xlNewDokumentZamAltInfo.GIDTyp} AND Atr_ObiNumer = {xlNewDokumentZamAltInfo.GIDNumer}  AND Atr_ObiLp = 0 AND Atr_ObiSubLp = 0 AND (AtK_Nazwa = 'Dzień Dostawy' )) begin 
                                    insert into cdn.Atrybuty(Atr_ObiTyp, Atr_ObiFirma, Atr_ObiNumer, Atr_ObiLp, Atr_ObiSubLp, Atr_AtkId, Atr_Wartosc, Atr_AtrTyp, Atr_AtrFirma, Atr_AtrNumer, Atr_AtrLp, Atr_AtrSubLp, Atr_OptimaId, Atr_Grupujacy, Atr_Pozycja) 
                                            VALUES
                                            ({xlNewDokumentZamAltInfo.GIDTyp},
                                                {xlNewDokumentZamAltInfo.GIDFirma},
                                                {xlNewDokumentZamAltInfo.GIDNumer},
                                            0, 
                                            0, (Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='Dzień Dostawy'),CAST(datediff(day, '18001228','{pickedDate.ToString("yyyyMMdd")}') AS VARCHAR), 0, 0, 0, 0, 0, 0, 0,1); end;";
                                string resp = ConnectToSQL(connection_string, zz_atr_new);
                                if (resp != "")
                                    LogAppendValue(Color.White, Color.DarkBlue, $@"Dodawanie atrybutu Dzień dostawy dla GID {xlNewDokumentZamAltInfo.GIDNumer} zwróciło {resp}");

                            }
                            LogAppendValue(Color.Black, Color.White, $@"Zamknięto dokument ZS o gid {xlNewDokumentZamAltInfo.GIDNumer} z kodem {rtrn_val} i trybem zamk {xLZamkniecieDokumentuInfo_20202.TrybZamkniecia}");
                        }
                    }
                    catch(Exception e)
                    {
                        MessageBox.Show("Błąd w readCSV else " + e.Message + "\n" + Subdeb.Key);
                    }
                }
            }

            
            pp.Dispose();
            pp.Close();

            //XLLogout();
            return records;
            }
        private List<newCSVList> ReadnewCSV(string fileName)
        {
             
            List<newCSVList> records = new List<newCSVList>();
            double cnt = 0;
            try
            {
                StreamReader sr = new StreamReader(fileName);
                records = File.ReadAllLines(fileName)
                               .Skip(1) 
                               .Where(s => !string.IsNullOrWhiteSpace(s))
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

            Pasek pp = new Pasek();
            pp.Show();
            pp.pBar1.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            pp.pBar1.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            pp.pBar1.Maximum = result.Count() + 1;
            // Set the initial value of the ProgressBar.
            pp.pBar1.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            pp.pBar1.Step = 1;
            LogAppendValue(Color.Black, Color.White,$@"Wybrano innego dostawcę ({delivery_name}) oraz tryb {RodzajDzialania}");
            if (RodzajDzialania != 1)
            {
                foreach (var Subclient in result)
                {
                    pp.pBar1.PerformStep();
                    pp.label2.Text = Subclient.Key;
                    if (Subclient.FirstOrDefault().TotalAmnt.Length < 1 || Subclient.FirstOrDefault().Barcode.Length < 1) continue;
                    LogAppendValue(Color.White, Color.LightBlue, $@"Subklient z pliku {Subclient.FirstOrDefault().Subclient} ma {Subclient.Count()} pozycji");
                    if (Subclient.Key != "")
                    {

                        query = $@"if not exists(select top 1 Atr_AtrNumer  from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId=AtK_ID where Atr_Wartosc='{Subclient.Key}' and AtK_Nazwa LIKE 'KOD %') 
                        select top 1 2 
                        else 
                        if exists(select top 1 ZaN_GIDNumer from cdn.ZamNag where ZaN_DokumentObcy = '{Subclient.First()}')
                        select top 1 1
                        else 
                        select top 1 0 ";
                        string msg = ConnectToSQL(connection_string, query);
                        if (msg == "2")
                        {
                            LogAppendValue(Color.White, Color.DarkRed,$"Błąd! brak klienta w bazie z atrybutem dla kodu {Subclient.Key}");
                            continue;
                        }
                        else if (msg == "1")
                        {
                            //richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,$@"Istnieje już taka faktura z numerem {Subclient.First().in}");
                            //continue;
                        }
                        
                        int[] ZZ_gid = new int[4] { 0, 0, 0, 0 };
                        string grupaPodatkowa = "A";
                        for (int OrderCount = 0; OrderCount < Subclient.Count(); OrderCount++)
                        {
                            int _barcodeCount = 0;
                            cnt = 0;
                            if (OrderCount == 0)
                            {
                                grupaPodatkowa = ConnectToSQL(connection_string, $@"select top 1 Knt_ExpoKraj from cdn.KntKarty where Knt_GIDNumer={ZZ_knt[0]}");
                                ZZ_gid = CreateZZ(ZZ_knt, Subclient.ElementAt(OrderCount).Grower, Subclient.ElementAt(OrderCount).Code, true, pickedDate, pickedDate, grupaPodatkowa);
                            } 
                            try
                            {
                                if (Double.TryParse(Subclient.ElementAt(OrderCount).TotalAmnt, out double temp) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                cnt += temp;
                                try
                                {
                                if (OrderCount < (Subclient.Count() - 1))
                                {
                                    while   (Subclient.ElementAt(OrderCount).Barcode == Subclient.ElementAt(OrderCount + 1).Barcode && 
                                            OrderCount < (Subclient.Count() - 1) &&
                                            Subclient.ElementAt(OrderCount).Price == Subclient.ElementAt(OrderCount + 1).Price &&
                                            Subclient.ElementAt(OrderCount).Article == Subclient.ElementAt(OrderCount + 1).Article)
                                        {
                                        _barcodeCount++;
                                        OrderCount++;
                                        if (Double.TryParse(Subclient.ElementAt(OrderCount).TotalAmnt, out double temp2) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                        cnt += temp2;
                                        if (OrderCount == (Subclient.Count() - 1))
                                        {
                                            LogAppendValue(Color.Black, Color.White, $@"Order count: {OrderCount} a Subdeb count {Subclient.Count()} więc robie break (ostatnia wartość duplikatu)");
                                            break;
                                        }
                                    }
                                }
                                }
                                catch (Exception ex2)
                                {
                                    LogAppendValue(Color.White, Color.DarkRed, $@"Bład {++errorCount} przy order count {OrderCount} i {Subclient.ElementAt(OrderCount).Article}  {ex2}");
                                }

                                if (cnt > 0)
                                {
                                    if (_barcodeCount>0) LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} Barcode {Subclient.ElementAt(OrderCount).Barcode} powtórzył się {_barcodeCount+1} razy");
                                    else LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} Barcode {Subclient.ElementAt(OrderCount).Barcode}");
                                    if(RodzajDzialania==2)
                                        AddNewElement_XL_ZZ(XLZam_id, Subclient.ElementAt(OrderCount),  cnt.ToString(), grupaPodatkowa, "ZZ");
                                    else
                                        AddNewElement_XL(XLZam_id, Subclient.ElementAt(OrderCount), AddNewIndex_XL(Subclient.ElementAt(OrderCount), pickedDate, RodzajDzialania, grupaPodatkowa), cnt.ToString(), grupaPodatkowa, "ZZ");
                                }
                                cnt = 0;
                            }
                            catch (Exception e)
                            {
                                throw new Exception("Błąd przy dodawaniu elementów do ZZ: " + e);
                            }
                        }
                        //XLZS_id 
                        CloseZZ(ZZ_knt, delivery_name, ZZ_gid, XLZam_id, pickedDate, pickedDate);
                        if (RodzajDzialania != 2) CreateZS(ZZ_gid, Subclient.Key, pickedDate, pickedDate);

                    }
                    else
                    {
                        /*
                       query = $@"
                       if exists(select top 1 ZaN_GIDNumer from cdn.ZamNag where ZaN_DokumentObcy = '{Subclient.First()}')
                       select top 1 1
                       else 
                       select top 1 0 ";
                       string msg = ConnectToSQL(connection_string, query);
                       if (msg == "1")
                       {
                           LogAppendValue(Color.White, Color.DarkRed, "Istnieje już taka faktura.");
                           continue;
                       }

                       try
                       {
                           //Zmienia się ZZ knt z powodu zamówienia ZZ.
                           ZZ_knt = ConnectToSQL(connection_string, $@"select CONCAT(Knt_GIDNumer,';', Knt_GIDTyp,';', Knt_GIDLp,';', Knt_GIDFirma) from cdn.KntKarty where Knt_Akronim=(select  top 1  r.Knt_Akronim 
                                   from cdn.AtrybutyKlasy as k 
                                   join cdn.Atrybuty as a on k.AtK_ID = Atr_AtkId 
                                   join cdn.KntKarty as r on a.Atr_ObiNumer = r.Knt_GIDNumer and a.Atr_ObiTyp = Knt_GIDTyp 
                                   where k.AtK_Nazwa = 'FormatImportuPlikuZZ' and a.Atr_Wartosc = '{wlasnyAkronim}' order by Knt_GIDNumer asc)")
                               .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                               .Select(x => int.Parse(x))
                               .ToArray();
                       }
                       catch (Exception e)
                       {
                           LogAppendValue(Color.White, Color.DarkRed,$"Wystąpił błąd nr {++errorCount} przy pobieraniu Dostawcy przy zamówieniu własnym ({wlasnyAkronim})  " + ZZ_knt[0] + " komunikat " + e);
                           continue;
                       }
                       */
                        int[] ZZ_gid    = new int[4] { 0, 0, 0, 0 };
                        int[] gid_knt   = new int[4] { 0, 0, 0, 0 };
                        string grupaPod = "A";

                        for (int OrderCount = 0; OrderCount < Subclient.Count(); OrderCount++)
                        {
                            cnt = 0;
                            int _barcodeCount = 0;
                            if (OrderCount == 0)
                            {
                                grupaPod = ConnectToSQL(connection_string, $@"select top 1 Knt_ExpoKraj from cdn.KntKarty where Knt_GIDNumer={ZZ_knt[0]}");

                                ZZ_gid = CreateZZ(ZZ_knt, Subclient.ElementAt(OrderCount).Grower, Subclient.ElementAt(OrderCount).Code, true, pickedDate, pickedDate, grupaPod);
                            }
                            //var UniquePositions = records.Where(z => z.Barcode == Subclient.ElementAt(OrderCount).Barcode).ToList();

                            try
                            {
                                if (Double.TryParse(Subclient.ElementAt(OrderCount).TotalAmnt, out double temp) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                cnt += temp;
                                if (OrderCount < Subclient.Count() - 1)
                                {
                                    while (Subclient.ElementAt(OrderCount).Barcode == Subclient.ElementAt(OrderCount + 1).Barcode &&
                                            OrderCount < (Subclient.Count() - 1) &&
                                            Subclient.ElementAt(OrderCount).Price == Subclient.ElementAt(OrderCount + 1).Price &&
                                            Subclient.ElementAt(OrderCount).Article == Subclient.ElementAt(OrderCount + 1).Article)
                                    {
                                        _barcodeCount++;
                                        OrderCount++;
                                        if (Double.TryParse(Subclient.ElementAt(OrderCount).TotalAmnt, out double temp2) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                        cnt += temp2;
                                    }
                                }
                                if (cnt > 0)
                                {
                                    if (_barcodeCount > 0) LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} Barcode {Subclient.ElementAt(OrderCount).Barcode} powtórzył się {_barcodeCount + 1} razy");
                                    else LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} Barcode {Subclient.ElementAt(OrderCount).Barcode}");
                                    if(RodzajDzialania==2)
                                        AddNewElement_XL_ZZ(XLZam_id, Subclient.ElementAt(OrderCount), cnt.ToString(), grupaPod, "ZZ'");
                                    else
                                        AddNewElement_XL(XLZam_id, Subclient.ElementAt(OrderCount), AddNewIndex_XL(Subclient.ElementAt(OrderCount), pickedDate, RodzajDzialania, grupaPod), cnt.ToString(), grupaPod, "ZZ'");
                                }
                                cnt = 0;
                            }
                            catch (Exception e)
                            {

                                throw new Exception("blad: " + e);
                            }

                        }
                        //XLZS_id 
                        int rtrn_val = CloseZZ(ZZ_knt, delivery_name, ZZ_gid, XLZam_id, pickedDate, pickedDate);
                        //LogAppendValue(Color.Black, Color.White,$@"Zamknięto dokument ZZ (zam. własne) z kodem {rtrn_val}");

                    }
                }
            }
            else
            {
                LogAppendValue(Color.DarkViolet, Color.White, $@"Wybrano innego dostawcę ({delivery_name}) oraz rodzaj działania=1 ");
                //AlternatywnyZS
                foreach (var Subclient in result)
                {
                    pp.pBar1.PerformStep();
                    pp.label2.Text = Subclient.Key;
                    if (Subclient.FirstOrDefault().Barcode.Length < 1) continue;
                    int rtrn_val = -1;
                    int _barcodeCount = 0;
                    int[] new_knt = new int[4] { 0, 0, 0, 0 };
                    cnt = 0;
                    XLDokumentZamNagInfo_20231 xlNewDokumentZamAltInfo = new XLDokumentZamNagInfo_20231();
                    for (int OrderCount = 0; OrderCount < Subclient.Count(); OrderCount++)
                    {
                        _barcodeCount = 0;
                        cnt = 0;
                        if (OrderCount == 0)
                        {
                            string odp = "";
                            try
                            {
                                try
                                {
                                    
                                    new_knt = ConnectToSQL(connection_string, $"if exists(select * from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId=AtK_ID where Atr_Wartosc='{(Subclient.ElementAt(OrderCount).Subclient == "" ? wlasnyAkronim : Subclient.ElementAt(OrderCount).Subclient)}' and AtK_Nazwa LIKE 'KOD %' )" +
                                       $" select top 1 CONCAT(Atr_ObiNumer,';', Atr_ObiTyp,';', Atr_ObiLp,';', Atr_ObiFirma)  from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId=AtK_ID where Atr_Wartosc='{(Subclient.ElementAt(OrderCount).Subclient == "" ? wlasnyAkronim : Subclient.ElementAt(OrderCount).Subclient)}' and AtK_Nazwa LIKE 'KOD %'" +
                                       $" else select top 1 CONCAT(Knt_GIDNumer,';',Knt_GIDTyp,';', Knt_GIDLp,';',Knt_GIDFirma)  from cdn.KntKarty  where KNt_Akronim='ZS BRAK ATRYBUTU' ")
                                       //+ $" AtK_Nazwa = 'KOD ' + '(select top 1 Knt_Nazwa1 from cdn.KntKarty where Knt_GIDNumer = {knt[0]} and Knt_GIDTyp = {knt[1]} and Knt_GIDLp = {knt[2]} and Knt_GIDFirma = {knt[3]})'")
                                       .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                       .Select(x => int.Parse(x))
                                       .ToArray();
                                    if (new_knt[0] <= 0)
                                    {
                                        LogAppendValue(Color.White, Color.DarkRed, $@"nie odnaleziono kontrahenta do ZS {new_knt[0]} !");
                                      
                                        break;
                                    }
                                    else LogAppendValue(Color.Black, Color.White, $@"dla KNT {(Subclient.ElementAt(OrderCount).Subclient == "" ? wlasnyAkronim : Subclient.ElementAt(OrderCount).Subclient)} odnaleziono gid {new_knt[0]}");
                                }
                                catch (Exception e)
                                {
                                    LogAppendValue(Color.White, Color.DarkRed, $@"nie odnaleziono kontrahenta do ZS (błąd split do array)  + {e}!");
                                     break;
                                } 
                                
                                xlNewDokumentZamAltInfo.Wersja = XLWersjaAPI;
                                xlNewDokumentZamAltInfo.Typ = 6; //ZS
                                xlNewDokumentZamAltInfo.Tryb = 2;
                                xlNewDokumentZamAltInfo.KntFirma = new_knt[3];
                                xlNewDokumentZamAltInfo.KntLp = new_knt[2];
                                xlNewDokumentZamAltInfo.KntNumer = new_knt[0];
                                xlNewDokumentZamAltInfo.KntTyp = new_knt[1];
                                //xlNewDokumentZamAltInfo.Rok = pickedDate.Year;
                                //xlNewDokumentZamAltInfo.Miesiac = pickedDate.Month;
                                //xlNewDokumentZamAltInfo.DataWystawienia = GetClarionDate(pickedDate);         // ewentualnie GetClarionDate(pickedDate)
                                //xlNewDokumentZamAltInfo.DataPotwierdzeniaOferty = GetClarionDate(pickedDate);
                                // DateTime.Parse(Subclient.FirstOrDefault().RUPHEID);
                                //xlNewDokumentZamAltInfo.DataPotwierdzeniaOferty = GetClarionDate(DateTime.Parse(Subclient.FirstOrDefault().RUPHEID));
                                // xlNewDokumentZamAltInfo.DataWystawienia = GetClarionDate(DateTime.Parse(Subclient.FirstOrDefault().RUPHEID));
                                xlNewDokumentZamAltInfo.DataPotwierdzeniaOferty = GetClarionDate(DateTime.Today);
                                xlNewDokumentZamAltInfo.DataWystawienia = GetClarionDate(DateTime.Today);
                                xlNewDokumentZamAltInfo.FlagaNB = "N";
                                xlNewDokumentZamAltInfo.ZamSeria = Subclient.ElementAt(OrderCount).Subclient != "" ? "DYSTR" : "CENTR";
                                xlNewDokumentZamAltInfo.Waluta = "PLN";

                                rtrn_val = cdn_api.cdn_api.XLNowyDokumentZam(sesja, ref AltZS_id, xlNewDokumentZamAltInfo);
                            }
                            catch (Exception e)
                            {
                                LogAppendValue(Color.White, Color.DarkRed,$@"Wystąpił błąd {++errorCount} dodawania ZS! Nie znaleziono kontrahenta!" + e + $@"a odp na wybór kontrahenta to {odp}");

                                XLZamkniecieDokumentuZamInfo_20231 xLZamkniecieDokumentuInfo_202022 = new XLZamkniecieDokumentuZamInfo_20231
                                {
                                    Wersja = XLWersjaAPI,
                                    TrybZamkniecia = 1,
                                };
                                break;
                            }

                            if (rtrn_val != 0 || new_knt[0] == 0)
                            {

                                if (rtrn_val != 33)
                                {
                                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount}. zakładania dokument ZS o gid {xlNewDokumentZamAltInfo.GIDNumer} z kodem {rtrn_val} i knt {xlNewDokumentZamAltInfo.KntNumer} oraz {odp}");
                                }
                                else
                                {
                                    LogAppendValue(Color.White, Color.DarkRed,
                                        $@"Błąd API {rtrn_val} : Brak licencji na moduł Zamówienia - prawdopodobnie potrzebne ponowne zalogowanie do XL");
                                }
                                XLZamkniecieDokumentuZamInfo_20231 xLZamkniecieDokumentuInfo_202022 = new XLZamkniecieDokumentuZamInfo_20231
                                {
                                    Wersja = XLWersjaAPI,
                                    TrybZamkniecia = 1,
                                };
                                break;
                            }
                            else LogAppendValue(Color.White, Color.DarkRed, $@"Utworzono ZS o GID {xlNewDokumentZamAltInfo.GIDNumer}"); 
                            LogAppendValue(Color.Black, Color.White, $@"OrderCount wynosi {OrderCount} a Sublcient count {Subclient.Count()}");
                        }
                        try
                        {
                            if (Double.TryParse(Subclient.ElementAt(OrderCount).TotalAmnt, out double temp) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                            cnt += temp;
                            if (OrderCount < (Subclient.Count() - 1))
                            {
                                while (Subclient.ElementAt(OrderCount).Barcode == Subclient.ElementAt(OrderCount + 1).Barcode &&
                                            OrderCount < (Subclient.Count() - 1) &&
                                            Subclient.ElementAt(OrderCount).Price == Subclient.ElementAt(OrderCount + 1).Price &&
                                            Subclient.ElementAt(OrderCount).Article == Subclient.ElementAt(OrderCount + 1).Article)
                                {
                                    _barcodeCount++;
                                    OrderCount++;
                                    if (Double.TryParse(Subclient.ElementAt(OrderCount).TotalAmnt, out double temp2) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd numer {++errorCount}Rzutowanie ilosci nie powiodło się");
                                    cnt += temp2;
                                    if (OrderCount == (Subclient.Count() - 1))
                                    {
                                        LogAppendValue(Color.Red, Color.White, $@"Order count: {OrderCount} a Subdeb count {Subclient.Count()}. Przerywam (ostatni duplikat)");
                                        break;
                                    }
                                }
                            }
                            if (cnt > 0)
                            {
                                if (_barcodeCount > 0) LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} Barcode {Subclient.ElementAt(OrderCount).Barcode} powtórzył się {_barcodeCount + 1} razy");
                                else LogAppendValue(Color.Green, Color.Black, $@"Pozycja nr {OrderCount} Barcode {Subclient.ElementAt(OrderCount).Barcode}");
                                AddNewElement_XL(AltZS_id, Subclient.ElementAt(OrderCount), AddNewIndex_XL(Subclient.ElementAt(OrderCount), pickedDate, RodzajDzialania,"A"), cnt.ToString(), "A", "ZS");
                            }
                            cnt = 0;
                        }
                        catch (Exception e)
                        {
                            rtrn_val = -1;
                            throw new Exception("blad: " + e);
                        } 
                    
                    }
                    if(rtrn_val!=-1)
                    { 
                    XLZamkniecieDokumentuZamInfo_20231 xLZamkniecieDokumentuInfo_20202 = new XLZamkniecieDokumentuZamInfo_20231
                    {
                        Wersja = XLWersjaAPI,
                        TrybZamkniecia = rtrn_val == 0 ? 0 : 1,
                    };
                    rtrn_val = cdn_api.cdn_api.XLZamknijDokumentZam(AltZS_id, xLZamkniecieDokumentuInfo_20202);

                    if(rtrn_val==0 && pickedDate!=DateTime.Today)
                        {
                            //ChangeDate(GetClarionDate(pickedDate), DateTime.Parse(Subclient.FirstOrDefault().RUPHEID), AltZS_id);
                            var str = ChangeDateZS(GetClarionDate(DateTime.Today), pickedDate, xlNewDokumentZamAltInfo.GIDNumer);

                            if (str != "")
                            {
                                LogAppendValue(Color.Black, Color.White, $@"Bład przy zmianie daty dostawy na {pickedDate}! {str}");
                            }
                            else LogAppendValue(Color.Black, Color.White, $@"Pomyślnie zmieniono date na {pickedDate}");

                        }
                        if (xLZamkniecieDokumentuInfo_20202.TrybZamkniecia == 0)
                        {

                            string zz_atr_new = $@"if not exists (select 1 from cdn.Atrybuty left join cdn.AtrybutyKlasy on AtK_ID=Atr_AtkId where Atr_ObiTyp = {xlNewDokumentZamAltInfo.GIDTyp} AND Atr_ObiNumer = {xlNewDokumentZamAltInfo.GIDNumer}  AND Atr_ObiLp = 0 AND Atr_ObiSubLp = 0 AND (AtK_Nazwa = 'Dzień Dostawy' )) begin 
                                    insert into cdn.Atrybuty(Atr_ObiTyp, Atr_ObiFirma, Atr_ObiNumer, Atr_ObiLp, Atr_ObiSubLp, Atr_AtkId, Atr_Wartosc, Atr_AtrTyp, Atr_AtrFirma, Atr_AtrNumer, Atr_AtrLp, Atr_AtrSubLp, Atr_OptimaId, Atr_Grupujacy, Atr_Pozycja) 
                                            VALUES
                                            ({xlNewDokumentZamAltInfo.GIDTyp},
                                                {xlNewDokumentZamAltInfo.GIDFirma},
                                                {xlNewDokumentZamAltInfo.GIDNumer},
                                            0, 
                                            0, (Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='Dzień Dostawy'),CAST(datediff(day, '18001228','{pickedDate.ToString("yyyyMMdd")}') AS VARCHAR), 0, 0, 0, 0, 0, 0, 0,1); 
                                    end;";
                            string odp = ConnectToSQL(connection_string, zz_atr_new);
                            if (odp != "")
                                LogAppendValue(Color.White, Color.DarkBlue, $@"Dodawanie atrybutu Dzien Dostawy do GID {xlNewDokumentZamAltInfo.GIDNumer} ma resp {odp}");
                        }
                        LogAppendValue(Color.Black, Color.White,$@"Zamknięto dokument ZS o gid {xlNewDokumentZamAltInfo.GIDNumer} z kodem {rtrn_val} i trybem zamk {xLZamkniecieDokumentuInfo_20202.TrybZamkniecia}");
                    }
                }
            }

            pp.Dispose();
            pp.Close();
            //richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,$@"Wylogowano! kod wylogowania {XLLogout()}!");
            return records;
        }

        private int AddNewElement_XL(int xLZam_id, newCSVList order, int[] vs, string TotalAmnt, string grupaPodatkowa, string typDok)
        {
             
            int str = -1;
            /*
            richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,
            "Nazwa: " + order.Article + Environment.NewLine +
            "TowarEAN: " + order.Barcode + Environment.NewLine +
            "Ilosc: " + order.TotalAmnt + Environment.NewLine +
            "TwrNumer: " + vs[0] + Environment.NewLine +
            "TwrTyp: " + vs[1] + Environment.NewLine +
            "TwrLp: " + vs[2] + Environment.NewLine +
            "TwrFirma: " + vs[3] + Environment.NewLine +
            "CenaOferowana: " + Regex.Replace(order.ClientPrice, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", ".") + Environment.NewLine +
            "CenaKatalogowa: " + Regex.Replace(order.Price, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", ".") + Environment.NewLine);
            */
            XLDokumentZamElemInfo_20231 xlDokumentZamInfo = new XLDokumentZamElemInfo_20231
            {
                Wersja = XLWersjaAPI,
                Nazwa = order.Article,
                TowarEAN = order.Barcode,
                Ilosc = TotalAmnt,
                TwrNumer = vs[0],
                TwrTyp = vs[1],
                TwrLp = vs[2],
                TwrFirma = vs[3],
                CenaOferowana = Regex.Replace(order.ClientPrice, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", "."),
                CenaKatalogowa = Regex.Replace(order.Price, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", "."),
                Rabat = "0",
                RabatReq=1
                //Wartosc= Regex.Replace(order.ClientPrice, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", ".")
            };
            if (typDok=="ZS")
            {
                xlDokumentZamInfo.Waluta="PLN";
                xlDokumentZamInfo.CenaOferowana = Regex.Replace(order.ClientPrice, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", ".");
                xlDokumentZamInfo.CenaKatalogowa = Regex.Replace(order.ClientPrice, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", ".");
            }
            if (grupaPodatkowa=="2")
            {
                xlDokumentZamInfo.GrupaVAT = "C";
                xlDokumentZamInfo.StawkaVAT = "0";
            }
            else if (typDok == "ZS")
            {
                var zsVat = ConnectToSQL(connection_string, $@"select Concat(Twr_StawkaPodSpr,';',Twr_GrupaPodSpr) from cdn.TwrKarty where Twr_GIDNumer={vs[0]}");
                try
                {
                    var zsTabVat = zsVat.Split(';');
                    xlDokumentZamInfo.StawkaVAT = zsTabVat[0];
                    xlDokumentZamInfo.GrupaVAT = zsTabVat[1];
                }
                catch (Exception ex)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Błąd nr {++errorCount} przy konwersji VAT dla {zsVat} ex: {ex}");
                }
            }
            if (RodzajDzialania == 1)
            {
                string temp = xlDokumentZamInfo.CenaOferowana;
                xlDokumentZamInfo.CenaOferowana = xlDokumentZamInfo.CenaKatalogowa;
                xlDokumentZamInfo.CenaKatalogowa = temp;
            }
            try
            {
                str = cdn_api.cdn_api.XLDodajPozycjeZam(xLZam_id, xlDokumentZamInfo);
            }
            catch (Exception e)
            {
                throw new Exception("Błąd nr " + e.ToString() + " kod API: " + str);
            }
            if (RodzajDzialania == 1)
            {
                try
                {
                    LogAppendValue(Color.White, Color.DarkBlue, "Dodawanie atrybutów");
                    XLAtrybutInfo_20231 atrybut1 = new XLAtrybutInfo_20231();
                    atrybut1.Wersja = XLWersjaAPI;
                    atrybut1.GIDFirma = xlDokumentZamInfo.GIDFirma;
                    atrybut1.GIDLp = xlDokumentZamInfo.GIDLp;
                    atrybut1.GIDNumer = xlDokumentZamInfo.GIDNumer;
                    atrybut1.GIDSubLp = 0;
                    atrybut1.GIDTyp = xlDokumentZamInfo.GIDTyp;
                    atrybut1.Wartosc = dostawa;
                    atrybut1.Klasa = "Nr dostawy";
                    int blad_atr = cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut1);

                    if (blad_atr != 0)
                    {
                        MessageBox.Show("Błąd atrybutu nr dostawy: " + blad_atr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    XLAtrybutInfo_20231 atrybut2 = new XLAtrybutInfo_20231();
                    atrybut2.Wersja = XLWersjaAPI;
                    atrybut2.GIDFirma = xlDokumentZamInfo.GIDFirma;
                    atrybut2.GIDLp = xlDokumentZamInfo.GIDLp;
                    atrybut2.GIDNumer = xlDokumentZamInfo.GIDNumer;
                    atrybut2.GIDSubLp = 0;
                    atrybut2.GIDTyp = xlDokumentZamInfo.GIDTyp;
                    atrybut2.Wartosc = "0";
                    atrybut2.Klasa = "Zużyta ilość";
                    blad_atr = cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut2);

                    if (blad_atr != 0)
                    {
                        MessageBox.Show("Błąd atrybutu nr dostawy: " + blad_atr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                   


                }
                catch (Exception e)
                {
                    MessageBox.Show("Błąd dodawania atrybutów do ZS: " + e.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            if (str != 0)
            {
                LogAppendValue(Color.White, Color.DarkRed, $@"Api przy dodaniu pozycji zwrocilo błąd (nr{++errorCount}) o nr {str}");
                XLKomunikatInfo_20231 info = new XLKomunikatInfo_20231();
                info.Wersja = XLWersjaAPI;
                info.Funkcja = 42;
                info.Blad = str;
                cdn_api.cdn_api.XLOpisBledu(info);
                LogAppendValue(Color.White, Color.DarkRed, $@"Opis błędu: {info.OpisBledu}");

            }
            else LogAppendValue(Color.Black, Color.White,$@"Dodano element do zamówienia o nazwie {xlDokumentZamInfo.Nazwa} w ilości {TotalAmnt}");
            return str;
        }

        private int AddNewElement_XL_ZZ(int xLZam_id, newCSVList order, string TotalAmnt, string grupaPodatkowa, string typDok)
        {

            int str = -1;
            int count =-1;
            int pozostalaIlosc = Convert.ToInt32(TotalAmnt);
            try
            {
                while (pozostalaIlosc > 0)
                {
                
                    int[] vs = AddNewIndex_XL_ZZ(order, pickedDate, RodzajDzialania, grupaPodatkowa);
                    count = vs.Count();
              
                    //Twr_GIDNumer , Twr_GIDTyp   , Twr_GIDLP  , Twr_GIDFirma ,ZaE_GIDNumer  , ZaE_GIDTyp   , ZaE_GIDLp  ,ZaE_GIDFirma, strybut IloscUzyta, ZaE_Ilosc 
                    int iloscPoz = vs[9] - vs[8];
                    if (iloscPoz > pozostalaIlosc || iloscPoz == 0)
                        iloscPoz = pozostalaIlosc;

                    XLDokumentZamElemInfo_20231 xlDokumentZamInfo = new XLDokumentZamElemInfo_20231
                    {
                        Wersja = XLWersjaAPI,
                        Nazwa = order.Article,
                        TowarEAN = order.Barcode,
                        Ilosc = Convert.ToString(iloscPoz),
                        TwrNumer = vs[0],
                        TwrTyp = vs[1],
                        TwrLp = vs[2],
                        TwrFirma = vs[3],
                        CenaOferowana = Regex.Replace(order.ClientPrice, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", "."),
                        CenaKatalogowa = Regex.Replace(order.Price, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", "."),
                        Rabat="0",
                        RabatReq=1
                    };

                    if (grupaPodatkowa == "2")
                    {
                        xlDokumentZamInfo.GrupaVAT = "C";
                        xlDokumentZamInfo.StawkaVAT = "0";
                    }

                    try
                    {
                        str = cdn_api.cdn_api.XLDodajPozycjeZam(xLZam_id, xlDokumentZamInfo);
                    }
                    catch (Exception e)
                    {
                        throw new Exception("Błąd nr " + e.ToString() + " kod API: " + str);
                    }

                    pozostalaIlosc -= iloscPoz;
                    //MessageBox.Show("Ilosc: "+ vs[9]+" \nZużyta ilość: "+vs[8]+"\nIlość na pozycje: "+iloscPoz+"\nPozostała ilość: " + pozostalaIlosc, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (vs[7] != 0)
                    {

                        string updateIlosc = "Update a set Atr_Wartosc = '" + Convert.ToString(iloscPoz+vs[8]) + "' FROM cdn.atrybuty as a join cdn.ZamElem as e on ZaE_GIDFirma = a.Atr_ObiFirma and ZaE_GIDLp = a.Atr_ObiLp and ZaE_GIDNumer = a.Atr_ObiNumer and ZaE_GIDTyp = a.Atr_ObiTyp "
                             + "WHERE Atr_AtkId = 51 and ZaE_GIDFirma = " + vs[7] + " and ZaE_GIDLp = " + vs[6] + " and ZaE_GIDNumer = " + vs[4] + " and ZaE_GIDTyp = " + vs[5];
                        Update(updateIlosc);

                    }

                    if (str != 0)
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Api przy dodaniu pozycji zwrocilo błąd (nr{++errorCount}) o nr {str}");
                        XLKomunikatInfo_20231 info = new XLKomunikatInfo_20231();
                        info.Wersja = XLWersjaAPI;
                        info.Funkcja = 42;
                        info.Blad = str;
                        cdn_api.cdn_api.XLOpisBledu(info);
                        LogAppendValue(Color.White, Color.DarkRed, $@"Opis błędu: {info.OpisBledu}");

                    }
                    else LogAppendValue(Color.Black, Color.White, $@"Dodano element do zamówienia o nazwie {xlDokumentZamInfo.Nazwa} w ilości {TotalAmnt}");
                
                }
            }
            catch (Exception e)
            {
                LogAppendValue(Color.White, Color.DarkRed, $@"Błąd przy dodaniu pozycji.  Długość tabeli {count} \nTreść błędu: {e.Message}");
            }
            return str;
        }

        private int[] AddNewIndex_XL(newCSVList order, DateTime pickedTime, int tryb, string grupaPodatkowa)
        {
            AttachThreadToClarion(true);

             
            int return_value = -1;
            bool is_tf = false;
            string qry_resp;
            string last_part = "";

            if(ConnectToSQL(connection_string, $"begin if exists(Select 1 from cdn.TwrKarty where Twr_Ean='{order.Barcode}' and Twr_Archiwalny=1) select top 1 1 else select top 1 0 end")!="0")
            {
                string updateEAN = $"UPDATE cdn.TwrKarty set Twr_EAN = '' WHERE Twr_Ean = '{order.Barcode}' and Twr_Archiwalny = 1";
                Update(updateEAN);
                LogAppendValue(Color.Black, Color.White, $@"Usunięto EAN {order.Barcode} o gid dla towarów archiwalnych");
            }

            if (ConnectToSQL(connection_string, $"begin if exists(Select 1 from cdn.TwrKarty where Twr_Ean='{order.Barcode}') select top 1 1 else select top 1 0 end") != "0")
            {
                try
                {
                    int[] knt_arr = ConnectListToSQL(connection_string, $@"SELECT TOP 1 Cast(Twr_GIDNumer as varchar) 
                      , Cast(Twr_GIDTyp as varchar) 
                      , Cast(Twr_GIDLP as varchar) 
                      , Cast(Twr_GIDFirma as varchar) 
                    FROM cdn.TwrKarty where Twr_Ean = '{order.Barcode}'").ToArray();
                    LogAppendValue(Color.Black, Color.White, $@"Odnaleziono towar dla EAN {order.Barcode} o gid {knt_arr[0]}");
                    if(RodzajDzialania==1)
                    {
                        string updateQuery = "update cdn.TwrKarty set Twr_EAN='' where Twr_GIDNumer="+knt_arr[0]+" and Twr_GIDTyp="+knt_arr[1]+" and Twr_GIDLp= "+knt_arr[2]+" and Twr_GIDFirma="+knt_arr[3];
                        Update(updateQuery);
                        LogAppendValue(Color.Black, Color.White, $@"Usunięto EAN {order.Barcode} o gid {knt_arr[0]}");
                    }
                    else
                        return knt_arr;
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed,$@"Wystąpił błąd rzutowania kontrahenta {e}");
                    return new int[4] { 0, 0, 0, 0 };
                }
            }
            string ogrodnik = "";
            if (order.Grower.Replace($@"'", $@"''").Length>1)
            {
                
            ogrodnik = ConnectToSQL(connection_string, $@"declare @resp as VARCHAR(MAX) 
            declare @grower as VARCHAR(MAX) 
            set @grower='{order.Grower.Replace($@"'", $@"''")}' 
            if not exists (select 1 from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower) 
                        begin 
			            insert into cdn.AtrybutyWartosci(AtW_AtKId, AtW_Wartosc, AtW_OddZrdId, AtW_Lp) 
                        Values( (select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2'), @grower, 0, 0); 
			            set @resp = Right('00000' + Convert(varchar(5),(select top 1 AtW_ID from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower)),5) 
			            end 
			            else 
			            begin 
                        set @resp = Right('00000' + Convert(varchar(5),(select top 1 AtW_ID from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower)),5) 
			            end select @resp");
                if (ogrodnik.Length > 10)
                {
                    LogAppendValue(Color.White, Color.DarkRed,$@"Wystąpił błąd przy dodawaniu id ogrodnika do atrybutu Hodowca2 {ogrodnik} !");
                    ogrodnik = "";
                }
            }
            string nazwa = $@"{order.VBN}-{order.S1}-{order.S2}-{order.S3}_{pickedTime.ToString("yyMMdd")}_{ogrodnik}";
            nazwa = Truncate(nazwa, 35);

            if (order.Subclient.Length == 0 || tryb == 2)
            {
                nazwa += "_TUFL";
                is_tf = true;
            }
            else
                nazwa += ("_" + order.Subclient);
            //Sprawdzic czy istnieje duplikat
            qry_resp = ConnectToSQL(connection_string, $@"declare @reszta_kodu varchar(max) 
                declare @vbn varchar(max) 
                declare @srodek varchar(max) 
                declare @count int 
                set @vbn = '{nazwa.Substring(0, nazwa.IndexOf('-'))}' 
                set @srodek = '' 
                set @reszta_kodu ='{nazwa.Substring(nazwa.IndexOf('-'))}' 
                set @count = 0 
                if not exists(Select 1 from cdn.TwrKarty where Twr_Kod=@vbn+@srodek+@reszta_kodu) 
	                begin 
	                select 0 
	                end 
                else 
	                begin 
	                while (Select 1 from cdn.TwrKarty where Twr_Kod=@vbn+@srodek+@reszta_kodu)=1  
	                begin 
	                set @count = @count+1 
	                set @srodek = '|'+CAST(@count as varchar(max)) 
	                end 
                select @count 
                end");
               
                if (qry_resp != "0" && int.TryParse(qry_resp, out int number) == true)
                {
                    LogAppendValue(Color.White, Color.DarkBlue, $@"Znaleziono duplikat nr {number} towaru {nazwa} o EAN {order.Barcode}");
                    try
                    {
                        nazwa = $@"{order.VBN}|{number}-{order.S1}-{order.S2}-{order.S3}_{pickedTime.ToString("yyMMdd")}_{ogrodnik}";
                        last_part = number.ToString();
                        nazwa = Truncate(nazwa, 35);
                        if (order.Subclient.Length == 0 || tryb == 2)
                        {
                            nazwa += "_TUFL";
                        }
                        else
                            nazwa += ("_" + order.Subclient);
                    }
                    catch (Exception ex)
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd {++errorCount} przy duplikacie {nazwa} z EAN {order.Barcode} {ex}!");
                        return new int[4] { 0, 0, 0, 0 };
                    }
                }
                else if (int.TryParse(qry_resp, out int numberek) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd {++errorCount} rzutowania dla Query {qry_resp}");
            LogAppendValue(Color.Black, Color.White, $@"Dodawanie nowego towaru o kodzie {nazwa} oraz nazwie {order.Article}");
            //Stawka podatkowa z kodu Code - jest tabelka w XL
            XLTowarInfo_20231 xLTowarInfo = new XLTowarInfo_20231
            {
                Wersja = XLWersjaAPI,
                GIDFirma = 0,
                GIDLp = 0,
                GIDNumer = 0,
                GIDTyp = 0,
                Typ = 1,
                Kod = $"{nazwa}", 
                StawkaPod = "23",
                GrupaPod = "A",
                StawkaPodSpr = "23",
                GrupaPodSpr = "A",
                Nazwa = order.Article,
                EAN = order.Barcode,
                //PCN = order.Code.PadLeft(8, '0'),
                //SWW = Regex.Replace(order.Code, ".{2}", "$0.").TrimEnd('.'),
                Jm = "szt.", 
            };
                ;
            //Mapowanie CN i PKWIU

            string grupaCN = "";
            try
            {
                string stawka = "23";
                var resp = ConnectListStrToSQL(connection_string, $@"IF EXISTS (select 1 from cdn.KodyCN where KCN_Kod='{order.Code.PadLeft(8, '0')}') 
                begin 
                select top 1 Cast(KCN_Kod as varchar), 
			                 Cast(KCN_StawkaVAT as varchar), 
			                 Cast(KCN_Opis as varchar(max)) 
			                 from cdn.KodyCN where KCN_Kod='{order.Code.PadLeft(8, '0')}' 
                end 
                else 
                select 'x', 'x', 'x'").ToArray();
                //MessageBox.Show($@"{resp[0]}, {resp[1]}, {resp[2]}");
                
                LogAppendValue(Color.Black, Color.White,$@"Dodawanie kodu KCN, stawka i opisu: {resp[0]}, {resp[1]}, {resp[2].Split('#').LastOrDefault()}");
                if (resp[0].Length > 1)
                {
                    xLTowarInfo.PCN = resp[0];
                    stawka = Regex.Replace(resp[1], @"[^\d]", "").TrimEnd('0');
                    switch (stawka)
                    {
                        case "23":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "A";
                            xLTowarInfo.GrupaPodSpr = "A";
                            break;
                        case "8":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "B";
                            xLTowarInfo.GrupaPodSpr = "B";
                            break;
                        case "5":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "G";
                            xLTowarInfo.GrupaPodSpr = "G";
                            break;
                        case "3":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "F";
                            xLTowarInfo.GrupaPodSpr = "F";
                            break;
                        case "0":

                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "C";
                            xLTowarInfo.GrupaPodSpr = "C";
                            break;
                        default:
                            xLTowarInfo.StawkaPod = "23";
                            xLTowarInfo.StawkaPodSpr = "23";
                            xLTowarInfo.GrupaPod = "A";
                            xLTowarInfo.GrupaPodSpr = "A";
                            break;
                    }
                    
                    if(resp[2].Contains("$"))
                    {
                        LogAppendValue(Color.DarkGreen, Color.White, $@"Grupa towarowa zawiera $: {resp[2]} " + "\n");
                        if (resp[2].Contains("#"))
                        {
                            grupaCN = resp[2].Substring(resp[2].IndexOf("$") + 1, resp[2].IndexOf("#")- resp[2].IndexOf("$")-1);
                            LogAppendValue(Color.DarkGreen, Color.White, $@"Grupa towarowa zawiera #, grupaCN: {grupaCN} " + "\n");
                        }
                        else
                        {
                            grupaCN = resp[2].Substring(resp[2].IndexOf("$") + 1);
                            LogAppendValue(Color.DarkGreen, Color.White, $@"Grupa towarowa nie zawiera #, grupaCN: {grupaCN} " + "\n");
                        }
                    

                    }
                    if (resp[2].Contains("#"))
                        xLTowarInfo.SWW = resp[2].Split('#').LastOrDefault();
                }
            }
            catch (Exception exce)
            {
                LogAppendValue(Color.White, Color.DarkRed,$@"Wystąpił błąd nr {++errorCount} konwersji podatkowej {exce} " + "\n");
            }
            if (grupaPodatkowa == "2")
                {
                    LogAppendValue(Color.Black, Color.White, $@"Towar o stawce podatkowej 0%!");
                    xLTowarInfo.StawkaPod = "0";
                    xLTowarInfo.GrupaPod = "C";
                }

            if (!is_tf)
            {
                xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE";
                LogAppendValue(Color.DarkGreen, Color.White, $@"Grupa towarowa: {xLTowarInfo.TwrGrupa} " + "\n");
            }
            else if(grupaCN!="")
            {
                //xLTowarInfo.TwrGrupa = grupaCN.TrimEnd();
                try
                {
                    var grupaQuery = ConnectToSQL(connection_string, "select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa='" + grupaCN.TrimEnd() + "'");
                    string[] completeGrupa = grupaQuery.Split(';');
                    if (completeGrupa.Length > 0)
                    {
                        xLTowarInfo.TwGNumer = Int32.Parse(completeGrupa[0]);
                        xLTowarInfo.TwGTyp = Int32.Parse(completeGrupa[1]);
                        xLTowarInfo.TwGLp = Int32.Parse(completeGrupa[2]);
                        xLTowarInfo.TwGFirma = Int32.Parse(completeGrupa[3]);
                    }
                    else LogAppendValue(Color.White, Color.DarkRed, $@"Błąd Nie znaleziono grupy dla towaru {xLTowarInfo.Nazwa}, odpowiedz to {grupaQuery}");
                    LogAppendValue(Color.DarkGreen, Color.White, $@"GrupaCN!="", Grupa towarowa: {xLTowarInfo.TwrGrupa} " + "\n");
                }
                catch(Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Błąd Nie znaleziono grupy dla towaru {xLTowarInfo.Nazwa}, treść błędu: {e.Message}");
                    xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE";
                }
            }

           else if (is_tf)
            {
                try
                { 
                    string group = FindGroup(order.Article.Replace($@"'", $@"''"));
                    LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $@"Odnaleziona nazwa grupy to {group}");
                    var resp = ConnectToSQL(GetConnectionString(), $@"IF EXISTS (select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa='{group}' and CDN.TwrGrupaPelnaNazwa(TwG_GIDNumer) like '%KWIATY CIĘTE/%') 
                    begin 
                    select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa='{group}' and CDN.TwrGrupaPelnaNazwa(TwG_GIDNumer) like '%KWIATY CIĘTE/%'
                    end 
                    else 
                    begin 
                    select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Kod='{group}' and CDN.TwrGrupaPelnaNazwa(TwG_GIDNumer) like '%KWIATY CIĘTE/%'
                    end");
                    LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $"Odnalezione GID to {resp}");

                    string[] completeGrupa = resp.Split(';');
                    if (completeGrupa.Length > 0 && group != "")
                    {
                        xLTowarInfo.TwGNumer = Int32.Parse(completeGrupa[0]);
                        xLTowarInfo.TwGTyp = Int32.Parse(completeGrupa[1]);
                        xLTowarInfo.TwGLp = Int32.Parse(completeGrupa[2]);
                        xLTowarInfo.TwGFirma = Int32.Parse(completeGrupa[3]);
                    }
                    else
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Błąd Nie znaleziono grupy dla towaru {xLTowarInfo.Nazwa}, odpowiedz to {resp}");
                        xLTowarInfo.TwGNumer = 26;
                        xLTowarInfo.TwGTyp = -16;
                        xLTowarInfo.TwGLp = 0;
                        xLTowarInfo.TwGFirma = 1403393;
                    }
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} dodawania grupy towarowej {e} " + "\n");
                    xLTowarInfo.TwrGrupa = "DODATKI";
                }
            }
            else xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE";
            LogAppendValue(Color.White, Color.DarkGreen, $@"Grupa towarowa: {xLTowarInfo.TwrGrupa} " + "\n");
            try
            {
                return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
               
            }
            catch (Exception e)
            {
                throw new Exception("Wystapil wyjatek przy dodawaniu towaru. Kod API: " + return_value + ", Wyjątek: " + e);

            }
            
            if (return_value != 0)
            {
                if(return_value!=1)LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} dodania towaru {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");
                LogAppendValue(Color.Black, Color.White, $"Poprawka wprowadzania grupy dla towaru {nazwa}");
                xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE"; 
                return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
                string group = ConnectToSQL(connection_string, "select top 1 CONCAT(TwG_GIDNumer, ';', TwG_GIDTyp, ';', TwG_GIDLp, ';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa = 'KWIATY DEDYKOWANE'");
                if(return_value!=0)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} poprawki x2 towaru <z nazwą> {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                    try
                    {
                        int[] spltGrp = group.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                           .Select(x => int.Parse(x))
                           .ToArray();
                        if (spltGrp.Length >= 3)
                        {
                            xLTowarInfo.TwGNumer = spltGrp[0];
                            xLTowarInfo.TwGTyp = spltGrp[1];
                            xLTowarInfo.TwGLp = spltGrp[2];
                            xLTowarInfo.TwGFirma = spltGrp[3];
                        }
                        return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
                        if (return_value != 0) LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} poprawki x3 towaru <z GID> {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                    }
                    catch (Exception)
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} pobierania GID grupy KWIATY DEDYKOWANE: {xLTowarInfo.Kod} z TwGNumer {xLTowarInfo.GIDNumer} TwGNumer {xLTowarInfo.TwGNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                    }
                }

                
            }
            if (RodzajDzialania == 1)
            {
                XLAtrybutInfo_20231 atrybut3 = new XLAtrybutInfo_20231();
                atrybut3.Wersja = XLWersjaAPI;
                atrybut3.GIDFirma = xLTowarInfo.GIDFirma;
                atrybut3.GIDLp = xLTowarInfo.GIDLp;
                atrybut3.GIDNumer = xLTowarInfo.GIDNumer;
                atrybut3.GIDSubLp = 0;
                atrybut3.GIDTyp = xLTowarInfo.GIDTyp;
                atrybut3.Wartosc = order.Barcode;
                atrybut3.Klasa = "Kod kreskowy";
                int blad_atr = cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut3);

                if (blad_atr != 0)
                {
                    MessageBox.Show("Błąd atrybutu kod kreskowy : " + blad_atr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            XLZamkniecieTowaruInfo_20231 xLZamkniecieTowaruInfo_20231 = new XLZamkniecieTowaruInfo_20231
            {
                Wersja = XLWersjaAPI,
                Tryb = return_value != 0 ? -1 : 0,
            };
            return_value = cdn_api.cdn_api.XLZamknijTowar(XLtowar_id, xLZamkniecieTowaruInfo_20231);
            Update("UPDATE  c SET TwC_TcnId = 2, [TwC_DataOd]=986562792 FROM [CDN].[TwrCeny] as c join cdn.TwrCenyNag on TwC_TcnId=TCN_Id WHERE TwC_TwrTyp = 16 AND TwC_TwrFirma = " + xLTowarInfo.GIDFirma + " AND TwC_TwrNumer = " + xLTowarInfo.GIDNumer + " and TCN_RodzajCeny=2");
            Update("UPDATE  c SET TwC_TcnId = 1, [TwC_DataOd]=986562792 FROM [CDN].[TwrCeny] as c join cdn.TwrCenyNag on TwC_TcnId=TCN_Id WHERE TwC_TwrTyp = 16 AND TwC_TwrFirma = " + xLTowarInfo.GIDFirma + " AND TwC_TwrNumer = " + xLTowarInfo.GIDNumer + " and TCN_RodzajCeny=1");

            if (return_value != 0)
            {
                LogAppendValue(Color.White, Color.DarkRed,$"Błąd {++errorCount} zamknięcia towaru: {xLTowarInfo.Kod} dodany do magazynu o gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");
            }
            else
            {

                LogAppendValue(Color.Black, Color.White,$@"Utworzono towar z nazwą {xLTowarInfo.Nazwa} o kodzie {xLTowarInfo.Kod}");
                if (is_tf == true)
                {
                    string resp = ConnectToSQL(connection_string, $@"update t 
                    set 
                    Twr_ESklep = 1 
                    from cdn.TwrKarty as t 
                    where t.Twr_GIDNumer = {xLTowarInfo.GIDNumer} and t.Twr_GIDTyp={xLTowarInfo.GIDTyp} and t.Twr_GIDFirma={xLTowarInfo.GIDFirma}");

                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do e-Sklepu {"\n"} {resp}");
                    }

                    resp = ConnectToSQL(connection_string, $@"update cdn.TwrAplikacje set TAP_Status=1, TAP_Dostepnosc=1, TAP_Flagi=66 ,
                    TAP_GrupaeSklep=(select top 1 TwG_GrONumer from cdn.TwrGrupy where TwG_GIDNumer={xLTowarInfo.GIDNumer} and TwG_GIDTyp ={xLTowarInfo.GIDTyp} order by TwG_SyncId desc)  where TAP_ObiNumer = {xLTowarInfo.GIDNumer} and TAP_ObiTyp={xLTowarInfo.GIDTyp}");
                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do TwrAplikacje {"\n"} {resp}"); 
                    }
                }
               /* else
                {
                    string resp = ConnectToSQL(connection_string, $@"update t 
                    set Twr_MobSpr = 1  
                    from cdn.TwrKarty as t 
                    where t.Twr_GIDNumer = {xLTowarInfo.GIDNumer} and t.Twr_GIDTyp={xLTowarInfo.GIDTyp} and t.Twr_GIDFirma={xLTowarInfo.GIDFirma}");

                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do e-Sklepu {"\n"} {resp}");
                    }
                }*/
                string msg = "";
                if (order.Subclient.Length != 0)
                {
                    msg = ConnectToSQL(connection_string, "update cdn.TwrCeny set TwC_Wartosc=Cast('" + Regex.Replace(order.Price, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", ".") + "' as decimal(18,2)), TwC_Marza='0.00' where TwC_TwrNumer = (select Twr_GIDNumer from cdn.TwrKarty where Twr_Kod = '" + xLTowarInfo.Kod + "') and TwC_TwrLp=1");
                    
                    if (msg != "") LogAppendValue(Color.White, Color.DarkRed,$"Błąd nr{++errorCount} przy aktualizacji ceny {msg}!");
                }
                msg = ConnectToSQL(connection_string, $@"update cdn.TwrKarty set TwR_EAN='{order.Barcode}' , Twr_MobSpr = 1  where Twr_GIDNumer ={xLTowarInfo.GIDNumer} and Twr_GIDtyp={xLTowarInfo.GIDTyp}");
                if (msg != "")
                {
                    LogAppendValue(Color.White, Color.DarkRed, $"Błąd przy aktualizacji kodu EAN ({order.Barcode}) (wiodące zera z przodu) {msg}!");
                }
                try
                {
                    string flowerType = ConnectToSQL(connection_string, $"select top 1 Typ from FlowerType where KodCN={order.Code.PadLeft(8, '0')}");
                    LogAppendValue(Color.Black, Color.White, $"Rodzaj towaru dla Kodu CN {order.Code.PadLeft(8, '0')} to  {flowerType}");
                    var atr_resp = AddAtribute(new string[8] { order.VBN, order.Grower, order.Subclient, order.S1, order.S2, order.S3, order.RUPHEID, order.ImgUrl }, new int[4] { xLTowarInfo.GIDFirma, xLTowarInfo.GIDLp, xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp }, flowerType);
                    if (atr_resp.Length > 0)
                    {
                        LogAppendValue(Color.White, Color.DarkBlue, $"Bład SQL przy dodawaniu atrybutów {atr_resp}");
                    }
                    else LogAppendValue(Color.Black, Color.White, $"Pomyślnie dodano atrybuty");
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $"Błąd {errorCount} Nie udało się załadować atrybutów " + e);
                }
                try 
                {
                    LogAppendValue( Color.Black, Color.White,"Dodawanie Obrazka");
                    if(order.ImgUrl.Length>0)
                    {
                        var img = DownloadImage(order.ImgUrl);
                        if (img != null)
                        {
                            UploadImage(xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp, ((img.Size.Width > 512 && img.Size.Height > 512) ? ImageToByte2(ResizeImage(img, 512, 512)) : ImageToByte2(img)), order.Article, order.VBN);

                        //Zmienianie Atrybutów
                        string atr_Resp = ConnectToSQL(connection_string,
                            $@"update cdn.DaneObiekty set DAO_MobSpr = 1 
                                where DAO_ObiNumer = (select Twr_GIDNumer from cdn.TwrKarty where Twr_Kod = '{xLTowarInfo.Kod}')");
                            if (atr_Resp!="")
                            {
                                LogAppendValue(Color.White, Color.DarkBlue, "Nie udało się dodać atrybutu DAO_mobspr, odp. SQL to " + atr_Resp);
                            }

                            atr_Resp = ConnectToSQL(connection_string,
                            $@"update k set DAB_MobSpr = 1 from cdn.DaneBinarne as k
                                join cdn.DaneObiekty on DAO_DABId=DAB_Id
                                where DAO_DABId=DAB_Id and DAO_ObiNumer = (select Twr_GIDNumer from cdn.TwrKarty where Twr_Kod = '{xLTowarInfo.Kod}')");
                            if (atr_Resp != "")
                            {
                                LogAppendValue(Color.White, Color.DarkBlue, "Nie udało się dodać atrybutu DAB_mobspr, odp. SQL to " + atr_Resp);
                            }
                        }
                        else LogAppendValue(Color.White, Color.DarkRed,"Obraz jest NULL");
                    }
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed,$"Bład {errorCount} Nie udało się załadować obrazu " + e);
                }
                if (ZZ_knt[0] > 0)
                {
                    string a = ConnectToSQL(connection_string, $@"declare @nip varchar(10)
                    set @nip = (select top 1 Knt_NipPrefiks from cdn.KntKarty where Knt_GIDNumer={ZZ_knt[0]} order by Knt_GIDNumer desc)
                    if @nip is not null 
                    begin 
                    update cdn.TwrKarty set Twr_KrajPoch=@nip where Twr_GIDNumer={xLTowarInfo.GIDNumer} and Twr_GIDLp={xLTowarInfo.GIDLp} and Twr_GIDTyp={xLTowarInfo.GIDTyp} and Twr_GIDFirma={xLTowarInfo.GIDFirma} end");
                    if (a != "") LogAppendValue(Color.White, Color.DarkRed,$@"Wystąpił błąd dodawania kraju {a}");
                }
            }
            return new int[4] { xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp, xLTowarInfo.GIDLp, xLTowarInfo.GIDFirma };
        }

        private int[] AddNewIndex_XL_ZZ(newCSVList order, DateTime pickedTime, int tryb, string grupaPodatkowa)
        {
            AttachThreadToClarion(true);


            int return_value = -1;
            bool is_tf = false;
            string qry_resp;
            string last_part = "";
            string czyIstnieje = ConnectToSQL(connection_string, "begin if exists(SELECT TOP(1) 1  FROM cdn.ZamElem as e "
                    + "join cdn.atrybuty as a1 on ZaE_GIDFirma = a1.Atr_ObiFirma and ZaE_GIDLp = a1.Atr_ObiLp and ZaE_GIDNumer = Atr_ObiNumer and ZaE_GIDTyp = a1.Atr_ObiTyp "
                    + "join cdn.atrybuty as a3 on ZaE_GIDFirma = a3.Atr_ObiFirma and ZaE_GIDLp = a3.Atr_ObiLp and ZaE_GIDNumer = a3.Atr_ObiNumer and ZaE_GIDTyp = a3.Atr_ObiTyp "
                    + "join cdn.TwrKarty as t on ZaE_TwrFirma = Twr_GIDFirma and e.ZaE_TwrLp = t.Twr_GIDLp and e.ZaE_TwrNumer = t.Twr_GIDNumer and e.ZaE_TwrTyp = t.Twr_GIDTyp "
                    + "join cdn.atrybuty as a2 on Twr_GIDFirma = a2.Atr_ObiFirma and Twr_GIDLp = a2.Atr_ObiLp and Twr_GIDNumer = a2.Atr_ObiNumer and Twr_GIDTyp = a2.Atr_ObiTyp "
                    + "where a1.Atr_AtkId = 50 and a1.Atr_Wartosc = '" + dostawa + "' and a3.Atr_AtkId = 51 and a3.Atr_Wartosc < e.ZaE_Ilosc and a2.Atr_AtkId = 52 and a2.Atr_Wartosc = '" + order.Barcode + "') select 1 else select 0 end");
            LogAppendValue(Color.Black, Color.White, $@"CzyIstnieje: { czyIstnieje}");
            if (czyIstnieje.Contains("1"))
            {
                try
                {
                    int[] knt_arr = ConnectListToSQLZZ(connection_string, $@"SELECT TOP(1) Cast(Twr_GIDNumer as varchar)  , Cast(Twr_GIDTyp as varchar)  , Cast(Twr_GIDLP as varchar)  , Cast(Twr_GIDFirma as varchar) ,
	                                Cast(ZaE_GIDNumer as varchar)  , Cast(ZaE_GIDTyp as varchar)  , Cast(ZaE_GIDLp as varchar)  , Cast(ZaE_GIDFirma as varchar) ,
	                                Cast(a3.atr_Wartosc as varchar), SUBSTRING(Cast(ZaE_Ilosc as varchar), 0, CHARINDEX('.', Cast(ZaE_Ilosc as varchar)))
                                    FROM cdn.ZamElem as e
                                    join cdn.atrybuty as a1 on ZaE_GIDFirma=a1.Atr_ObiFirma and ZaE_GIDLp=a1.Atr_ObiLp and ZaE_GIDNumer=Atr_ObiNumer and ZaE_GIDTyp=a1.Atr_ObiTyp
                                    join cdn.atrybuty as a3 on ZaE_GIDFirma=a3.Atr_ObiFirma and ZaE_GIDLp=a3.Atr_ObiLp and ZaE_GIDNumer=a3.Atr_ObiNumer and ZaE_GIDTyp=a3.Atr_ObiTyp
                                    join cdn.TwrKarty as t on ZaE_TwrFirma=Twr_GIDFirma and e.ZaE_TwrLp=t.Twr_GIDLp and e.ZaE_TwrNumer=t.Twr_GIDNumer and e.ZaE_TwrTyp=t.Twr_GIDTyp
                                    join cdn.atrybuty as a2 on Twr_GIDFirma=a2.Atr_ObiFirma and Twr_GIDLp=a2.Atr_ObiLp and Twr_GIDNumer=a2.Atr_ObiNumer and Twr_GIDTyp=a2.Atr_ObiTyp
                                    where a1.Atr_AtkId=50 and a1.Atr_Wartosc='{dostawa}' and a3.Atr_AtkId=51 and a3.Atr_Wartosc<e.ZaE_Ilosc and a2.Atr_AtkId=52 and a2.Atr_Wartosc='{order.Barcode}'
                                    ORDER BY ZaE_GIDNumer").ToArray();
                    LogAppendValue(Color.Black, Color.White, $@"Odnaleziono towar dla EAN {order.Barcode} o gid {knt_arr[0]} dla dostawy {dostawa}, użyta ilość: { knt_arr[8]}");
                    //MessageBox.Show($@"Odnaleziono towar dla EAN {order.Barcode} o gid {knt_arr[0]}");


                    return knt_arr;
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd rzutowania kontrahenta {e}");
                    return new int[10] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                }
            }

            if (ConnectToSQL(connection_string, $"begin if exists(Select 1 from cdn.TwrKarty where Twr_Ean='{order.Barcode}' and Twr_Archiwalny=1) select top 1 1 else select top 1 0 end") != "0")
            {
                string updateEAN = $"UPDATE cdn.TwrKarty set Twr_EAN = '' WHERE Twr_Ean = '{order.Barcode}' and Twr_Archiwalny = 1";
                Update(updateEAN);
                LogAppendValue(Color.Black, Color.White, $@"Usunięto EAN {order.Barcode} o gid dla towarów archiwalnych");
            }

            if (ConnectToSQL(connection_string, $"begin if exists(Select 1 from cdn.TwrKarty where Twr_Ean='{order.Barcode}') select top 1 '1' else select top 1 '0' end") != "0")
            {
                try
                {
                    int[] knt_arr = ConnectListToSQLZZ(connection_string, $@"SELECT TOP 1 Cast(Twr_GIDNumer as varchar) 
                      , Cast(Twr_GIDTyp as varchar) 
                      , Cast(Twr_GIDLP as varchar) 
                      , Cast(Twr_GIDFirma as varchar), '0', '0', '0', '0', '0', '0'
                    FROM cdn.TwrKarty where Twr_Ean = '{order.Barcode}'").ToArray();
                    LogAppendValue(Color.Black, Color.White, $@"Odnaleziono towar dla EAN {order.Barcode} o gid {knt_arr[0]}");

                    if (ConnectToSQL(connection_string, "SELECT Twr_Kod From cdn.TwrKarty WHERE Twr_GIDNumer=" + knt_arr[0] + " and Twr_GIDTyp=" + knt_arr[1] + " and Twr_GIDLp= " + knt_arr[2] + " and Twr_GIDFirma=" + knt_arr[3]).Contains("TUFL"))
                        return knt_arr;
                    else
                    {
                        string updateQuery = "update cdn.TwrKarty set Twr_EAN='' where Twr_GIDNumer=" + knt_arr[0] + " and Twr_GIDTyp=" + knt_arr[1] + " and Twr_GIDLp= " + knt_arr[2] + " and Twr_GIDFirma=" + knt_arr[3];
                        Update(updateQuery);
                        LogAppendValue(Color.Black, Color.White, $@"Usunięto EAN {order.Barcode} o gid {knt_arr[0]}");
                    }

                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd usuwania EAN {e}");
                    return new int[10] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                }
            }

            string ogrodnik = "";
            if (order.Grower.Replace($@"'", $@"''").Length > 1)
            {

                ogrodnik = ConnectToSQL(connection_string, $@"declare @resp as VARCHAR(MAX) 
            declare @grower as VARCHAR(MAX) 
            set @grower='{order.Grower.Replace($@"'", $@"''")}' 
            if not exists (select 1 from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower) 
                        begin 
			            insert into cdn.AtrybutyWartosci(AtW_AtKId, AtW_Wartosc, AtW_OddZrdId, AtW_Lp) 
                        Values( (select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2'), @grower, 0, 0); 
			            set @resp = Right('00000' + Convert(varchar(5),(select top 1 AtW_ID from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower)),5) 
			            end 
			            else 
			            begin 
                        set @resp = Right('00000' + Convert(varchar(5),(select top 1 AtW_ID from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower)),5) 
			            end select @resp");
                if (ogrodnik.Length > 10)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd przy dodawaniu id ogrodnika do atrybutu Hodowca2 {ogrodnik} !");
                    ogrodnik = "";
                }
            }
            string nazwa = $@"{order.VBN}-{order.S1}-{order.S2}-{order.S3}_{pickedTime.ToString("yyMMdd")}_{ogrodnik}";
            nazwa = Truncate(nazwa, 35);

            if (order.Subclient.Length == 0 || tryb == 2)
            {
                nazwa += "_TUFL";
                is_tf = true;
            }
            else
                nazwa += ("_" + order.Subclient);
            //Sprawdzic czy istnieje duplikat
            qry_resp = ConnectToSQL(connection_string, $@"declare @reszta_kodu varchar(max) 
                declare @vbn varchar(max) 
                declare @srodek varchar(max) 
                declare @count int 
                set @vbn = '{nazwa.Substring(0, nazwa.IndexOf('-'))}' 
                set @srodek = '' 
                set @reszta_kodu ='{nazwa.Substring(nazwa.IndexOf('-'))}' 
                set @count = 0 
                if not exists(Select 1 from cdn.TwrKarty where Twr_Kod=@vbn+@srodek+@reszta_kodu) 
	                begin 
	                select 0 
	                end 
                else 
	                begin 
	                while (Select 1 from cdn.TwrKarty where Twr_Kod=@vbn+@srodek+@reszta_kodu)=1  
	                begin 
	                set @count = @count+1 
	                set @srodek = '|'+CAST(@count as varchar(max)) 
	                end 
                select @count 
                end");

            if (qry_resp != "0" && int.TryParse(qry_resp, out int number) == true)
            {
                LogAppendValue(Color.White, Color.DarkBlue, $@"Znaleziono duplikat nr {number} towaru {nazwa} o EAN {order.Barcode}");
                try
                {
                    nazwa = $@"{order.VBN}|{number}-{order.S1}-{order.S2}-{order.S3}_{pickedTime.ToString("yyMMdd")}_{ogrodnik}";
                    last_part = number.ToString();
                    nazwa = Truncate(nazwa, 35);
                    if (order.Subclient.Length == 0 || tryb == 2)
                    {
                        nazwa += "_TUFL";
                    }
                    else
                        nazwa += ("_" + order.Subclient);
                }
                catch (Exception ex)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd {++errorCount} przy duplikacie {nazwa} z EAN {order.Barcode} {ex}!");
                    return new int[10] { 0, 0, 0, 0 , 0, 0, 0, 0, 0, 0 };
                }
            }
            else if (int.TryParse(qry_resp, out int numberek) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd {++errorCount} rzutowania dla Query {qry_resp}");
            LogAppendValue(Color.Black, Color.White, $@"Dodawanie nowego towaru o kodzie {nazwa} oraz nazwie {order.Article}");
            //Stawka podatkowa z kodu Code - jest tabelka w XL
            XLTowarInfo_20231 xLTowarInfo = new XLTowarInfo_20231
            {
                Wersja = XLWersjaAPI,
                GIDFirma = 0,
                GIDLp = 0,
                GIDNumer = 0,
                GIDTyp = 0,
                Typ = 1,
                Kod = $"{nazwa}",
                StawkaPod = "23",
                GrupaPod = "A",
                StawkaPodSpr = "23",
                GrupaPodSpr = "A",
                Nazwa = order.Article,
                EAN = order.Barcode,
                //PCN = order.Code.PadLeft(8, '0'),
                //SWW = Regex.Replace(order.Code, ".{2}", "$0.").TrimEnd('.'),
                Jm = "szt.",
            };
            ;
            //Mapowanie CN i PKWIU

            string grupaCN = "";
            try
            {
                string stawka = "23";
                var resp = ConnectListStrToSQL(connection_string, $@"IF EXISTS (select 1 from cdn.KodyCN where KCN_Kod='{order.Code.PadLeft(8, '0')}') 
                begin 
                select top 1 Cast(KCN_Kod as varchar), 
			                 Cast(KCN_StawkaVAT as varchar), 
			                 Cast(KCN_Opis as varchar(max)) 
			                 from cdn.KodyCN where KCN_Kod='{order.Code.PadLeft(8, '0')}' 
                end 
                else 
                select 'x', 'x', 'x'").ToArray();
                //MessageBox.Show($@"{resp[0]}, {resp[1]}, {resp[2]}");

                LogAppendValue(Color.Black, Color.White, $@"Dodawanie kodu KCN, stawka i opisu: {resp[0]}, {resp[1]}, {resp[2].Split('#').LastOrDefault()}");
                if (resp[0].Length > 1)
                {
                    xLTowarInfo.PCN = resp[0];
                    stawka = Regex.Replace(resp[1], @"[^\d]", "").TrimEnd('0');
                    switch (stawka)
                    {
                        case "23":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "A";
                            xLTowarInfo.GrupaPodSpr = "A";
                            break;
                        case "8":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "B";
                            xLTowarInfo.GrupaPodSpr = "B";
                            break;
                        case "5":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "G";
                            xLTowarInfo.GrupaPodSpr = "G";
                            break;
                        case "3":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "F";
                            xLTowarInfo.GrupaPodSpr = "F";
                            break;
                        case "0":

                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "C";
                            xLTowarInfo.GrupaPodSpr = "C";
                            break;
                        default:
                            xLTowarInfo.StawkaPod = "23";
                            xLTowarInfo.StawkaPodSpr = "23";
                            xLTowarInfo.GrupaPod = "A";
                            xLTowarInfo.GrupaPodSpr = "A";
                            break;
                    }

                    if (resp[2].Contains("$"))
                    {
                        LogAppendValue(Color.DarkGreen, Color.White, $@"Grupa towarowa zawiera $: {resp[2]} " + "\n");
                        if (resp[2].Contains("#"))
                        {
                            grupaCN = resp[2].Substring(resp[2].IndexOf("$") + 1, resp[2].IndexOf("#") - resp[2].IndexOf("$") - 1);
                            LogAppendValue(Color.DarkGreen, Color.White, $@"Grupa towarowa zawiera #, grupaCN: {grupaCN} " + "\n");
                        }
                        else
                        {
                            grupaCN = resp[2].Substring(resp[2].IndexOf("$") + 1);
                            LogAppendValue(Color.DarkGreen, Color.White, $@"Grupa towarowa nie zawiera #, grupaCN: {grupaCN} " + "\n");
                        }


                    }
                    if (resp[2].Contains("#"))
                        xLTowarInfo.SWW = resp[2].Split('#').LastOrDefault();
                }
            }
            catch (Exception exce)
            {
                LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} konwersji podatkowej {exce} " + "\n");
            }
            if (grupaPodatkowa == "2")
            {
                LogAppendValue(Color.Black, Color.White, $@"Towar o stawce podatkowej 0%!");
                xLTowarInfo.StawkaPod = "0";
                xLTowarInfo.GrupaPod = "C";
            }

            if (!is_tf)
            {
                xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE";
                LogAppendValue(Color.DarkGreen, Color.White, $@"Grupa towarowa: {xLTowarInfo.TwrGrupa} " + "\n");
            }
            else if (grupaCN != "")
            {
                try
                {
                    var grupaQuery = ConnectToSQL(connection_string, "select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa='" + grupaCN.TrimEnd() + "'");
                    string[] completeGrupa = grupaQuery.Split(';');
                    if (completeGrupa.Length > 0)
                    {
                        xLTowarInfo.TwGNumer = Int32.Parse(completeGrupa[0]);
                        xLTowarInfo.TwGTyp = Int32.Parse(completeGrupa[1]);
                        xLTowarInfo.TwGLp = Int32.Parse(completeGrupa[2]);
                        xLTowarInfo.TwGFirma = Int32.Parse(completeGrupa[3]);
                    }
                    else LogAppendValue(Color.White, Color.DarkRed, $@"Błąd Nie znaleziono grupy dla towaru {xLTowarInfo.Nazwa}, odpowiedz to {grupaQuery}");
                    LogAppendValue(Color.DarkGreen, Color.White, $@"GrupaCN!="", Grupa towarowa: {xLTowarInfo.TwrGrupa} " + "\n");
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Błąd Nie znaleziono grupy dla towaru {xLTowarInfo.Nazwa}, treść błędu: {e.Message}");
                    xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE";
                }
            }

            else if (is_tf)
            {
                try
                {
                    string group = FindGroup(order.Article.Replace($@"'", $@"''"));
                    LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $@"Odnaleziona nazwa grupy to {group}");
                    var resp = ConnectToSQL(GetConnectionString(), $@"IF EXISTS (select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa='{group}' and CDN.TwrGrupaPelnaNazwa(TwG_GIDNumer) like '%KWIATY CIĘTE/%') 
                    begin 
                    select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa='{group}' and CDN.TwrGrupaPelnaNazwa(TwG_GIDNumer) like '%KWIATY CIĘTE/%'
                    end 
                    else 
                    begin 
                    select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Kod='{group}' and CDN.TwrGrupaPelnaNazwa(TwG_GIDNumer) like '%KWIATY CIĘTE/%'
                    end");
                    LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $"Odnalezione GID to {resp}");

                    string[] completeGrupa = resp.Split(';');
                    if (completeGrupa.Length > 0 && group != "")
                    {
                        xLTowarInfo.TwGNumer = Int32.Parse(completeGrupa[0]);
                        xLTowarInfo.TwGTyp = Int32.Parse(completeGrupa[1]);
                        xLTowarInfo.TwGLp = Int32.Parse(completeGrupa[2]);
                        xLTowarInfo.TwGFirma = Int32.Parse(completeGrupa[3]);
                    }
                    else
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Błąd Nie znaleziono grupy dla towaru {xLTowarInfo.Nazwa}, odpowiedz to {resp}");
                        xLTowarInfo.TwGNumer = 26;
                        xLTowarInfo.TwGTyp = -16;
                        xLTowarInfo.TwGLp = 0;
                        xLTowarInfo.TwGFirma = 1403393;
                    }
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} dodawania grupy towarowej {e} " + "\n");
                    xLTowarInfo.TwrGrupa = "DODATKI";
                }
            }
            else xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE";
            LogAppendValue(Color.White, Color.DarkGreen, $@"Grupa towarowa: {xLTowarInfo.TwrGrupa} " + "\n");
            try
            {
                return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
               

            }
            catch (Exception e)
            {
                throw new Exception("Wystapil wyjatek przy dodawaniu towaru. Kod API: " + return_value + ", Wyjątek: " + e);

            }

            if (return_value != 0)
            {
                if (return_value != 1) LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} dodania towaru {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");
                LogAppendValue(Color.Black, Color.White, $"Poprawka wprowadzania grupy dla towaru {nazwa}");
                xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE";
                return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
                string group = ConnectToSQL(connection_string, "select top 1 CONCAT(TwG_GIDNumer, ';', TwG_GIDTyp, ';', TwG_GIDLp, ';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa = 'KWIATY DEDYKOWANE'");
                if (return_value != 0)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} poprawki x2 towaru <z nazwą> {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                    try
                    {
                        int[] spltGrp = group.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                           .Select(x => int.Parse(x))
                           .ToArray();
                        if (spltGrp.Length >= 3)
                        {
                            xLTowarInfo.TwGNumer = spltGrp[0];
                            xLTowarInfo.TwGTyp = spltGrp[1];
                            xLTowarInfo.TwGLp = spltGrp[2];
                            xLTowarInfo.TwGFirma = spltGrp[3];
                        }
                        return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
                        if (return_value != 0) LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} poprawki x3 towaru <z GID> {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                    }
                    catch (Exception)
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} pobierania GID grupy KWIATY DEDYKOWANE: {xLTowarInfo.Kod} z TwGNumer {xLTowarInfo.GIDNumer} TwGNumer {xLTowarInfo.TwGNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                    }
                }


            }
           /* if (RodzajDzialania == 1)
            {
                XLAtrybutInfo_20231 atrybut3 = new XLAtrybutInfo_20231();
                atrybut3.Wersja = XLWersjaAPI;
                atrybut3.GIDFirma = xLTowarInfo.GIDFirma;
                atrybut3.GIDLp = xLTowarInfo.GIDLp;
                atrybut3.GIDNumer = xLTowarInfo.GIDNumer;
                atrybut3.GIDSubLp = 0;
                atrybut3.GIDTyp = xLTowarInfo.GIDTyp;
                atrybut3.Wartosc = order.Barcode;
                atrybut3.Klasa = "Kod kreskowy";
                int blad_atr = cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut3);

                if (blad_atr != 0)
                {
                    MessageBox.Show("Błąd atrybutu kod kreskowy : " + blad_atr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }*/

            XLZamkniecieTowaruInfo_20231 xLZamkniecieTowaruInfo_20231 = new XLZamkniecieTowaruInfo_20231
            {
                Wersja = XLWersjaAPI,
                Tryb = return_value != 0 ? -1 : 0,
            };
            return_value = cdn_api.cdn_api.XLZamknijTowar(XLtowar_id, xLZamkniecieTowaruInfo_20231);
            Update("UPDATE  c SET TwC_TcnId = 2, [TwC_DataOd]=986562792 FROM [CDN].[TwrCeny] as c join cdn.TwrCenyNag on TwC_TcnId=TCN_Id WHERE TwC_TwrTyp = 16 AND TwC_TwrFirma = " + xLTowarInfo.GIDFirma + " AND TwC_TwrNumer = " + xLTowarInfo.GIDNumer + " and TCN_RodzajCeny=2");
            Update("UPDATE  c SET TwC_TcnId = 1, [TwC_DataOd]=986562792 FROM [CDN].[TwrCeny] as c join cdn.TwrCenyNag on TwC_TcnId=TCN_Id WHERE TwC_TwrTyp = 16 AND TwC_TwrFirma = " + xLTowarInfo.GIDFirma + " AND TwC_TwrNumer = " + xLTowarInfo.GIDNumer + " and TCN_RodzajCeny=1");

            if (return_value != 0)
            {
                LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} zamknięcia towaru: {xLTowarInfo.Kod} dodany do magazynu o gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");
            }
            else
            {

                LogAppendValue(Color.Black, Color.White, $@"Utworzono towar z nazwą {xLTowarInfo.Nazwa} o kodzie {xLTowarInfo.Kod}");
                if (is_tf == true)
                {
                    string resp = ConnectToSQL(connection_string, $@"update t 
                    set  Twr_ESklep = 1 
                    from cdn.TwrKarty as t 
                    where t.Twr_GIDNumer = {xLTowarInfo.GIDNumer} and t.Twr_GIDTyp={xLTowarInfo.GIDTyp} and t.Twr_GIDFirma={xLTowarInfo.GIDFirma}");

                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do e-Sklepu {"\n"} {resp}");
                    }

                    resp = ConnectToSQL(connection_string, $@"update cdn.TwrAplikacje set TAP_Status=1, TAP_Dostepnosc=1, TAP_Flagi=66 , 
                    TAP_GrupaeSklep=(select top 1 TwG_GrONumer from cdn.TwrGrupy where TwG_GIDNumer={xLTowarInfo.GIDNumer} and TwG_GIDTyp ={xLTowarInfo.GIDTyp} order by TwG_SyncId desc)  where TAP_ObiNumer = {xLTowarInfo.GIDNumer} and TAP_ObiTyp={xLTowarInfo.GIDTyp}");
                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do TwrAplikacje {"\n"} {resp}");
                    }
                }
               /* else
                {
                    string resp = ConnectToSQL(connection_string, $@"update t 
                    set Twr_MobSpr = 1  
                    from cdn.TwrKarty as t 
                    where t.Twr_GIDNumer = {xLTowarInfo.GIDNumer} and t.Twr_GIDTyp={xLTowarInfo.GIDTyp} and t.Twr_GIDFirma={xLTowarInfo.GIDFirma}");

                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do e-Sklepu {"\n"} {resp}");
                    }
                }*/
                string msg = "";
                if (order.Subclient.Length != 0)
                {
                    msg = ConnectToSQL(connection_string, "update cdn.TwrCeny set TwC_Wartosc=Cast('" + Regex.Replace(order.Price, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", ".") + "' as decimal(18,2)), TwC_Marza='0.00' where TwC_TwrNumer = (select Twr_GIDNumer from cdn.TwrKarty where Twr_Kod = '" + xLTowarInfo.Kod + "') and TwC_TwrLp=1");

                    if (msg != "") LogAppendValue(Color.White, Color.DarkRed, $"Błąd nr{++errorCount} przy aktualizacji ceny {msg}!");
                }
                msg = ConnectToSQL(connection_string, $@"update cdn.TwrKarty set TwR_EAN='{order.Barcode}' , Twr_MobSpr = 1 where Twr_GIDNumer ={xLTowarInfo.GIDNumer} and Twr_GIDtyp={xLTowarInfo.GIDTyp}");
                if (msg != "")
                {
                    LogAppendValue(Color.White, Color.DarkRed, $"Błąd przy aktualizacji kodu EAN ({order.Barcode}) (wiodące zera z przodu) {msg}!");
                }
                try
                {
                    string flowerType = ConnectToSQL(connection_string, $"select top 1 Typ from FlowerType where KodCN={order.Code.PadLeft(8, '0')}");
                    LogAppendValue(Color.Black, Color.White, $"Rodzaj towaru dla Kodu CN {order.Code.PadLeft(8, '0')} to  {flowerType}");
                    var atr_resp = AddAtribute(new string[8] { order.VBN, order.Grower, order.Subclient, order.S1, order.S2, order.S3, order.RUPHEID, order.ImgUrl }, new int[4] { xLTowarInfo.GIDFirma, xLTowarInfo.GIDLp, xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp }, flowerType);
                    if (atr_resp.Length > 0)
                    {
                        LogAppendValue(Color.White, Color.DarkBlue, $"Bład SQL przy dodawaniu atrybutów {atr_resp}");
                    }
                    else LogAppendValue(Color.Black, Color.White, $"Pomyślnie dodano atrybuty");
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $"Błąd {errorCount} Nie udało się załadować atrybutów " + e);
                }
                try
                {
                    LogAppendValue(Color.Black, Color.White, "Dodawanie Obrazka");
                    if (order.ImgUrl.Length > 0)
                    {
                        var img = DownloadImage(order.ImgUrl);
                        if (img != null)
                        {
                            UploadImage(xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp, ((img.Size.Width > 512 && img.Size.Height > 512) ? ImageToByte2(ResizeImage(img, 512, 512)) : ImageToByte2(img)), order.Article, order.VBN);

                            //Zmienianie Atrybutów
                            string atr_Resp = ConnectToSQL(connection_string,
                                $@"update cdn.DaneObiekty set DAO_MobSpr = 1 
                                where DAO_ObiNumer = (select Twr_GIDNumer from cdn.TwrKarty where Twr_Kod = '{xLTowarInfo.Kod}')");
                            if (atr_Resp != "")
                            {
                                LogAppendValue(Color.White, Color.DarkBlue, "Nie udało się dodać atrybutu DAO_mobspr, odp. SQL to " + atr_Resp);
                            }

                            atr_Resp = ConnectToSQL(connection_string,
                           $@"update k set DAB_MobSpr = 1 from cdn.DaneBinarne as k
                                join cdn.DaneObiekty on DAO_DABId=DAB_Id
                                where DAO_DABId=DAB_Id and DAO_ObiNumer = (select Twr_GIDNumer from cdn.TwrKarty where Twr_Kod = '{xLTowarInfo.Kod}')");
                            if (atr_Resp != "")
                            {
                                LogAppendValue(Color.White, Color.DarkBlue, "Nie udało się dodać atrybutu DAB_mobspr, odp. SQL to " + atr_Resp);
                            }

                        }
                        else LogAppendValue(Color.White, Color.DarkRed, "Obraz jest NULL");
                    }
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $"Bład {errorCount} Nie udało się załadować obrazu " + e);
                }
                if (ZZ_knt[0] > 0)
                {
                    string a = ConnectToSQL(connection_string, $@"declare @nip varchar(10)
                    set @nip = (select top 1 Knt_NipPrefiks from cdn.KntKarty where Knt_GIDNumer={ZZ_knt[0]} order by Knt_GIDNumer desc)
                    if @nip is not null 
                    begin 
                    update cdn.TwrKarty set Twr_KrajPoch=@nip where Twr_GIDNumer={xLTowarInfo.GIDNumer} and Twr_GIDLp={xLTowarInfo.GIDLp} and Twr_GIDTyp={xLTowarInfo.GIDTyp} and Twr_GIDFirma={xLTowarInfo.GIDFirma} end");
                    if (a != "") LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd dodawania kraju {a}");
                }
            }
            return new int[10] { xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp, xLTowarInfo.GIDLp, xLTowarInfo.GIDFirma, 0, 0, 0, 0, 0, 0 };
        }
        private int[] CreateZZ(int[] knt, string dostawca, string dok_obcy, bool CzySeria, DateTime czasDost, DateTime wybranaData, string expokraj)
        {
            try
            {
                //string dostawca = "duyvenvoorde";
                int return_value = -1;
                AttachThreadToClarion(true);
                //(Knt_GIDNumer,';', Knt_GIDTyp,';', Knt_GIDLp,';', Knt_GIDFirma)
                XLDokumentZamNagInfo_20231 xlDokumentZamInfo = new XLDokumentZamNagInfo_20231
                {
                    Wersja = XLWersjaAPI,
                    Typ = 5, //ZZ
                    Tryb = 2,
                    KntNumer = knt[0],
                    KntFirma = knt[3],
                    KntLp = knt[2],
                    KntTyp = knt[1],
                    Waluta = "EUR",
                    FlagaNB = "N",
                    //Miesiac = wybranaData.Month,
                    //Rok = wybranaData.Year,
                    DataWystawienia = GetClarionDate(czasDost),         // ewentualnie GetClarionDate(pickedDate)
                    DataPotwierdzeniaOferty = GetClarionDate(czasDost), // ewentualnie GetClarionDate(pickedDate)
                    ZamSeria = CzySeria == true ? "DYSTR" : "CENTR",
                    //DokumentObcy = dok_obcy,
                    //Opis = dok_obcy,
                    ExpoNorm = -1
                };
                if (expokraj == "2") xlDokumentZamInfo.ExpoNorm = 10;
                var kursResp = ConnectToSQL(connection_string,
                                $@"declare @time date 
                            declare @diff int 
                            set @time = (select Convert(date,GetDate()-1)) 
                            set @diff = (select DATEDIFF(s,'1990-01-01',@time)) 
                            select top 1 Concat(Cast(WaE_KursL as int),';', WaE_KursM) from cdn.WalElem where WaE_Symbol='EUR' and WaE_KursTS<@diff order by WaE_kursTS desc");
                if (kursResp.IndexOf(';') > 0)
                {
                    string[] KursLM = kursResp.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                    try
                    {
                        xlDokumentZamInfo.KursL = Int32.Parse(KursLM[0]);
                        xlDokumentZamInfo.KursM = Int32.Parse(KursLM[1]);
                    }
                    catch (Exception)
                    {

                    }
                }
                try
                {
                    return_value = cdn_api.cdn_api.XLNowyDokumentZam(sesja, ref XLZam_id, xlDokumentZamInfo);

                }
                catch (Exception e)
                {

                    throw new Exception("Wystapil wyjatek przy dodawaniu nowego dokumentu ZZ. Kod API: " + return_value + ", Wyjątek: " + e);

                }
                if (return_value != 0)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Błąd o nr {++errorCount}! Wystapil wyjatek przy dodawaniu nowego dokumentu ZZ. Kod API: " + return_value);
                    if (return_value == 33)
                        MessageBox.Show("Brak licencji na moduł XL Zamówienia. Sprawdź dane na kluczu", "Komunikat");
                }
                else LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $@"Rozpoczęcie dodawania ZZ o gid: {xlDokumentZamInfo.GIDNumer} z KnT_GIDNumer: {knt[0]}");

                return new int[4] { xlDokumentZamInfo.GIDFirma, xlDokumentZamInfo.GIDLp, xlDokumentZamInfo.GIDNumer, xlDokumentZamInfo.GIDTyp };
            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd podczas dodawania dokumentu ZZ: " + e.Message, "Komunikat");
                return new int[4] { 0, 0, 0, 0 };
            }

            
            
        }
        private int CloseZZ(int[] knt, string dostawca, int[] ZZ_Gid, int zamID,DateTime czasDostawy, DateTime wybranaData)
        {
            AttachThreadToClarion(true);
             
            int rtnr_value = -1;
            XLZamkniecieDokumentuZamInfo_20231 xLZamkniecieDokumentu = new XLZamkniecieDokumentuZamInfo_20231
            {
                Wersja = XLWersjaAPI,
                TrybZamkniecia = ZZ_Gid[2] != 0 ? 2 : 1,
            };
            rtnr_value = cdn_api.cdn_api.XLZamknijDokumentZam(zamID, xLZamkniecieDokumentu);
            if (rtnr_value != 0)
            {
                LogAppendValue(Color.White, Color.DarkRed,
                    $@"Błąd nr {++errorCount}dodawania dokumentu ZZ dla kontrahenta {
                        _ = (knt[0]>0?(ConnectToSQL(connection_string, $@"select top 1 Knt_Akronim from cdn.KntKarty where Knt_GIDNumer={knt[0]}")):"Brak kontrahenta")}; Kod API {rtnr_value}");
            } else
            {
                {
                    if(czasDostawy!=wybranaData)
                    {
                        var str=ChangeDate(GetClarionDate(czasDostawy), pickedDate, ZZ_Gid[2]);

                        if (str != "")
                        {
                            LogAppendValue(Color.Black, Color.White, $@"Bład przy zmianie daty dostawy na {czasDostawy}! {str}");
                        }
                        else LogAppendValue(Color.Black, Color.White, $@"Pomyślnie zmieniono date na {czasDostawy}");
                            

                    }


                    string zz_Atr = $@"if not exists (select 1 from cdn.Atrybuty left join cdn.AtrybutyKlasy on AtK_ID=Atr_AtkId where Atr_ObiTyp = {ZZ_Gid[3]} AND Atr_ObiNumer = {ZZ_Gid[2]}  AND Atr_ObiLp = 0 AND Atr_ObiSubLp = 0 AND (AtK_Nazwa = '{"Źródło"}' 
                    OR Atr_AtkID = (Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='{"Zamówienie ślubne"}') )) begin 
                    insert into cdn.Atrybuty(Atr_ObiTyp, Atr_ObiFirma, Atr_ObiNumer, Atr_ObiLp, Atr_ObiSubLp, Atr_AtkId, Atr_Wartosc, Atr_AtrTyp, Atr_AtrFirma, Atr_AtrNumer, Atr_AtrLp, Atr_AtrSubLp, Atr_OptimaId, Atr_Grupujacy, Atr_Pozycja) 
                                    VALUES({ZZ_Gid[3].ToString()},
                                    {ZZ_Gid[0].ToString()},
                                    {ZZ_Gid[2].ToString()},
                                    0,
                                    0,(Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='{"Źródło"}'),'{"Sklep Holandia"}', 0, 0, 0, 0, 0, 0, 0,1) 
                 ,({ZZ_Gid[3].ToString()},
                   {ZZ_Gid[0].ToString()},
                   {ZZ_Gid[2].ToString()},
                   0, 0, (Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='{"Zamówienie ślubne"}'),'NIE', 0, 0, 0, 0, 0, 0, 0,1)
                 ,({ZZ_Gid[3].ToString()},
                   {ZZ_Gid[0].ToString()},
                   {ZZ_Gid[2].ToString()},
                   0, 0, 41,CAST(datediff(day, '18001228','{wybranaData.ToString("yyyyMMdd")}') AS VARCHAR), 0, 0, 0, 0, 0, 0, 0,1)
                ; end";

                    ConnectToSQL(connection_string, zz_Atr);
                    LogAppendValue(Color.Black, Color.White,$@"Zamknięto dokument ZZ dla kontrahenta {ConnectToSQL(connection_string, $@"select top 1 Knt_Akronim from cdn.KntKarty where Knt_GIDNumer={knt[0]}")} o KnT_GIDNumer={knt[0]}; Nazwa dokumentu to: " + ConnectToSQL(connection_string, $@"select top 1 cdn.NazwaObiektu({ZZ_Gid[3]},{ZZ_Gid[2]},0,2)"));
                }
            }

            return rtnr_value;
        }
        private int CreateZS(int[] ZZinfo, string subdeb, DateTime czasDostawy, DateTime wybranaData)
        {
            AttachThreadToClarion(true);
             
            int rtrn_val = -1;
            int[] new_knt = new int[4] { 0, 0, 0, 0 };

            try
            {
                new_knt = ConnectToSQL(connection_string, $"select top 1 CONCAT(Atr_ObiNumer,';', Atr_ObiTyp,';', Atr_ObiLp,';', Atr_ObiFirma)  from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId=AtK_ID where Atr_Wartosc='{(subdeb == "" ? wlasnyAkronim : subdeb)}' and AtK_Nazwa LIKE 'KOD %'")
                   //+ $" AtK_Nazwa = 'KOD ' + '(select top 1 Knt_Nazwa1 from cdn.KntKarty where Knt_GIDNumer = {knt[0]} and Knt_GIDTyp = {knt[1]} and Knt_GIDLp = {knt[2]} and Knt_GIDFirma = {knt[3]})'")
                   .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                   .Select(x => int.Parse(x))
                   .ToArray();
                if (new_knt[0] <= 0)
                {
                    LogAppendValue(Color.White, Color.DarkRed,$@"nie odnaleziono kontrahenta {(subdeb == "" ? wlasnyAkronim : subdeb)} do ZS {new_knt[0]} !");
                    return -1;
                } else LogAppendValue(Color.Black, Color.White, $@"dla KNT {(subdeb == "" ? wlasnyAkronim : subdeb)} odnaleziono gid {new_knt[0]}");
            }
            catch (Exception e)
            {
                LogAppendValue(Color.White, Color.DarkRed, $@"nie odnaleziono kontrahenta {(subdeb == "" ? wlasnyAkronim : subdeb)} do ZS (błąd split do array) + {e}!");
                return -1;
            }

            XLDokumentZamNagInfo_20231 xlNewDokumentZamInfo = new XLDokumentZamNagInfo_20231
            {
                Wersja = XLWersjaAPI,
                Typ = 6, //ZS
                Tryb = 2,
                KntFirma = new_knt[3],
                KntLp = new_knt[2],
                KntNumer = new_knt[0],
                KntTyp = new_knt[1],
                //Rok = wybranaData.Year,
                //Miesiac = wybranaData.Month,
                DataPotwierdzeniaOferty = GetClarionDate(czasDostawy),
                DataWystawienia = GetClarionDate(czasDostawy),
                FlagaNB = "N",
                Waluta = "PLN",
                ZamSeria = subdeb != "" ? "DYSTR" : "CENTR"
            };
            
            
            rtrn_val = cdn_api.cdn_api.XLNowyDokumentZam(sesja, ref XLZS_id, xlNewDokumentZamInfo);
            if (rtrn_val!=0)
            {
                if (rtrn_val!=33)
                {
                    LogAppendValue(Color.White, Color.DarkRed,
                        $@"Błąd dodawania ZS o gid {xlNewDokumentZamInfo.GIDNumer} z błędem nr {++errorCount} i kodem api {rtrn_val}");
                }
                else
                {
                    LogAppendValue(Color.White, Color.DarkRed,
                        $@"Błąd API {rtrn_val} : Brak licencji na moduł Zamówienia - prawdopodobnie potrzebne ponowne zalogowanie do XL");
                }
                XLZamkniecieDokumentuZamInfo_20231 xLZamkniecieDokumentuInfo2_20202 = new XLZamkniecieDokumentuZamInfo_20231
                {
                    Wersja = XLWersjaAPI,
                    TrybZamkniecia = rtrn_val == 0 ? 0 : 1,
                };
                rtrn_val = cdn_api.cdn_api.XLZamknijDokumentZam(XLZS_id, xLZamkniecieDokumentuInfo2_20202);
                return rtrn_val;
            } else LogAppendValue(Color.Black, Color.White,$@"Rozpoczęcie dodawania ZS o gid {xlNewDokumentZamInfo.GIDNumer} z KnT_GIDNumer={new_knt[0]}");
            new_knt = ConnectToSQL(connection_string, $@"if exists(select * from cdn.ZamNag where ZaN_GIDFirma = {ZZinfo[0]} and ZaN_GIDLp = {ZZinfo[1]} and ZaN_GIDNumer = {ZZinfo[2]} and ZaN_GIDTyp = {ZZinfo[3]}) 
                select CONCAT(ZaN_KntNumer,';',ZaN_KntTyp,';',ZaN_KntFirma,';',Zan_KntLp) from cdn.ZamNag where ZaN_GIDFirma = {ZZinfo[0]} and ZaN_GIDLp = {ZZinfo[1]} and ZaN_GIDNumer = {ZZinfo[2]} and ZaN_GIDTyp = {ZZinfo[3]} 
                else select '0;0;0;0'").Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
           .Select(x => int.Parse(x))
           .ToArray();
            string Twr_elem = ConnectToSQL(connection_string, $@"
            declare @maxCount integer 
            declare @Count integer 
            declare @txt varchar(max) 
            declare @gidnr integer 
            declare @kntnr integer 
            set @kntnr = {new_knt[0]}  
            set @gidnr = {ZZinfo[2]}  
            set @count = 1 
            set @txt = '' 
            set @maxCount = (select top 1 Count(ZaE_GIDNumer) from cdn.ZamElem where ZaE_KntNumer=@kntnr and ZaE_TwrKod!='A-VISTA' 
				              and ZaE_GIDNumer= @gidnr 
				              group by ZaE_GIDNumer) 
            while @count<=@maxCount 
	            begin 
	            if @txt !='' set @txt = @txt+'#'
		        set @txt = @txt+(select e.val from ( 
				            select Concat(ZaE_TwrTyp,';', 
                            ZaE_TwrFirma,';', 
                            ZaE_TwrLp,';', 
                            ZaE_TwrNumer,';', 
                            ZaE_Ilosc,';', 
                            ZaE_TwrNazwa,';', 
                            ZaE_StawkaPod,';', 
                            ZaE_GrupaPod,';', 
                            ZaE_CenaUzgodniona,';',
                            ZaE_CenaOferowana,';') as val, ROW_NUMBER() OVER (order by ZaE_TwrNumer asc) as rownumber from cdn.ZamElem 
                            where ZaE_KntNumer=@kntnr and ZaE_TwrKod!='A-VISTA' 
                            and ZaE_GIDNumer=@gidnr)
							as e where e.rownumber=@Count) 
							 +
							(select Concat(Twr_StawkaPodSpr,';',Twr_GrupaPodSpr) from cdn.TwrKarty where Twr_GIDNumer=
								(select e.val from 
									(select ZaE_TwrNumer as val, ROW_NUMBER() OVER (order by ZaE_TwrNumer asc) as rownumber from cdn.ZamElem 
									where ZaE_KntNumer=@kntnr and ZaE_TwrKod!='A-VISTA' 
									and ZaE_GIDNumer=@gidnr) 
								as e where e.rownumber=@Count) 
							) 
		            set @Count=@Count+1 
		            END
            Select @txt as txt");
            LogAppendValue(Color.White, Color.LightBlue,"Towary info to: " +Twr_elem);
             //else
            {

                try
                {
                    string[] Twr_elem_duo = Twr_elem.Split(new[] { '#' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var el in Twr_elem_duo)
                    {
                        string[] index = el.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                        if (index.Length >= 6)
                        {
                            XLDokumentZamElemInfo_20231 xlDokumentZamElemInfo = new XLDokumentZamElemInfo_20231
                            {
                                Wersja = XLWersjaAPI,
                                TwrTyp = Int32.Parse(index[0]),
                                TwrFirma = Int32.Parse(index[1]),
                                TwrLp = Int32.Parse(index[2]),
                                TwrNumer = Int32.Parse(index[3]),
                                Ilosc = index[4],
                                Nazwa = index[5],
                                StawkaVAT = index[10],
                                GrupaVAT = index[11],
                                CenaKatalogowa = index[9],
                                CenaOferowana = index[9],
                                Waluta = "PLN"
                            };

                            rtrn_val = cdn_api.cdn_api.XLDodajPozycjeZam(XLZS_id, xlDokumentZamElemInfo);
                            if (rtrn_val != 0)
                            {
                                LogAppendValue(Color.White, Color.DarkRed,$@"Błąd {++errorCount} dodawania elementu zam o twrgid: " + index[3] + " return value " + rtrn_val);
                            }
                        }

                    }
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, "Błąd nr " + e.ToString() + " kod API: " + rtrn_val);
                    throw new Exception("Błąd nr " + e.ToString() + " kod API: " + rtrn_val);
                }
            }

            if (rtrn_val != 0)
            {
                LogAppendValue(Color.White, Color.DarkRed,$@"Wystąpił błąd {++errorCount} zakładania dokument ZS o gid {ZamNag.ZaN_GIDNumer} z kodem {rtrn_val} i knt {ZamNag.ZaN_KntNumer}");
            }
            XLZamkniecieDokumentuZamInfo_20231 xLZamkniecieDokumentuInfo_20202 = new XLZamkniecieDokumentuZamInfo_20231
            {
                Wersja = XLWersjaAPI,
                TrybZamkniecia = rtrn_val == 0 ? 0 : 1,
            };
            rtrn_val = cdn_api.cdn_api.XLZamknijDokumentZam(XLZS_id, xLZamkniecieDokumentuInfo_20202);
            if (rtrn_val==0)
            {
                if (czasDostawy != wybranaData)
                {
                    var str = ChangeDate(GetClarionDate(czasDostawy), pickedDate, xlNewDokumentZamInfo.GIDNumer);

                    if (str != "")
                    {
                        LogAppendValue(Color.Black, Color.White, $@"Bład przy zmianie daty dostawy na {czasDostawy}! {str}");
                    }
                    else LogAppendValue(Color.Black, Color.White, $@"Pomyślnie zmieniono date na {czasDostawy}");


                }
            }
            string zz_atr_new = $@"if not exists (select 1 from cdn.Atrybuty left join cdn.AtrybutyKlasy on AtK_ID=Atr_AtkId where Atr_ObiTyp = {xlNewDokumentZamInfo.GIDTyp} AND Atr_ObiNumer = {xlNewDokumentZamInfo.GIDNumer}  AND Atr_ObiLp = 0 AND Atr_ObiSubLp = 0 AND (AtK_Nazwa = 'ZrodloZS' )) begin 
                insert into cdn.Atrybuty(Atr_ObiTyp, Atr_ObiFirma, Atr_ObiNumer, Atr_ObiLp, Atr_ObiSubLp, Atr_AtkId, Atr_Wartosc, Atr_AtrTyp, Atr_AtrFirma, Atr_AtrNumer, Atr_AtrLp, Atr_AtrSubLp, Atr_OptimaId, Atr_Grupujacy, Atr_Pozycja) 
                                VALUES({xlNewDokumentZamInfo.GIDTyp},
                                    {xlNewDokumentZamInfo.GIDFirma},
                                    {xlNewDokumentZamInfo.GIDNumer},
                                0, 
                                0, (Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa = 'ZrodloZS'), (select cdn.NazwaObiektu({ZZinfo[3]},{ZZinfo[2]},0,2)), {ZZinfo[3]}, {ZZinfo[0]}, {ZZinfo[2]}, 0, 0, 0, 0, 1),
                                ({xlNewDokumentZamInfo.GIDTyp},
                                    {xlNewDokumentZamInfo.GIDFirma},
                                    {xlNewDokumentZamInfo.GIDNumer},
                                0, 
                                0, (Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='Dzień Dostawy'),CAST(datediff(day, '18001228','{wybranaData.ToString("yyyyMMdd")}') AS VARCHAR), 0, 0, 0, 0, 0, 0, 0,1)
                ; end;";
            ConnectToSQL(connection_string, zz_atr_new);
            //richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,$@"Zamknięto dokument ZS o gid {ZamNag.ZaN_GIDNumer} z kodem {rtrn_val} i trybem zamk {xLZamkniecieDokumentuInfo_20202.TrybZamkniecia}");
            LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black,$@"Zamknięto dokument ZS dla konntrahenta o akronimie {ConnectToSQL(connection_string, $@"select top 1 Knt_Akronim from cdn.KntKarty where Knt_GIDNumer={xlNewDokumentZamInfo.KntNumer}")} i KnT_GIDNumber={new_knt[0]}. Nazwa dokumentu to: " + ConnectToSQL(connection_string, $@"select top 1 cdn.NazwaObiektu({xlNewDokumentZamInfo.GIDTyp},{xlNewDokumentZamInfo.GIDNumer},0,2)") + $@"z kodem {rtrn_val} i trybem zamk {xLZamkniecieDokumentuInfo_20202.TrybZamkniecia}");
            return rtrn_val;
        }
        public int AddElement_XL(int IDDokumentuZam, CSVList csvlist, int[] index, string ilosc, string grupaPod, string typDok)
        {

            int str = -1;

            XLDokumentZamElemInfo_20231 xlDokumentZamInfo = new XLDokumentZamElemInfo_20231
            {
                Wersja = XLWersjaAPI,
                Towar = csvlist.Articlegroup,
                Nazwa = csvlist.Article,
                TowarEAN = csvlist.Barcode,
                Ilosc = ilosc,
                TwrNumer = index[0],
                TwrTyp = index[1],
                TwrLp = index[2],
                TwrFirma = index[3],
                CenaOferowana = Regex.Replace(csvlist.ClientVal, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", "."),
                CenaKatalogowa = Regex.Replace(csvlist.Price, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", "."),
                Rabat = "0",
                RabatReq=1
            };
            if (typDok == "ZS")
            {
                xlDokumentZamInfo.Waluta = "PLN";
                xlDokumentZamInfo.CenaOferowana = Regex.Replace(csvlist.ClientVal, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", ".");
                xlDokumentZamInfo.CenaKatalogowa = Regex.Replace(csvlist.ClientVal, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", ".");
            }
            if (RodzajDzialania == 1)
            {
                string temp = xlDokumentZamInfo.CenaOferowana;
                xlDokumentZamInfo.CenaOferowana = xlDokumentZamInfo.CenaKatalogowa;
                xlDokumentZamInfo.CenaKatalogowa = temp;
            }
            if (grupaPod=="2")
            {
                xlDokumentZamInfo.GrupaVAT = "C";
                xlDokumentZamInfo.StawkaVAT = "0";
            }
            else if (typDok == "ZS")
            {
                var zsVat = ConnectToSQL(connection_string, $@"select Concat(Twr_StawkaPodSpr,';',Twr_GrupaPodSpr) from cdn.TwrKarty where Twr_GIDNumer={index[0]}");
                try
                {
                    var zsTabVat = zsVat.Split(';');
                    xlDokumentZamInfo.StawkaVAT = zsTabVat[0];
                    xlDokumentZamInfo.GrupaVAT = zsTabVat[1];
                }
                catch (Exception ex)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Błąd nr {++errorCount} przy konwersji VAT dla {zsVat} ex: {ex}");
                }  
            }
            try
            {
                str = cdn_api.cdn_api.XLDodajPozycjeZam(IDDokumentuZam, xlDokumentZamInfo);
            }
            catch (Exception e)
            {
                throw new Exception("Błąd nr " + e.ToString() + " kod API: " + str);
            }

            //LogAppendValue(Color.White, Color.DarkBlue, "Rodzaj działania: "+RodzajDzialania);
            if (RodzajDzialania==1)
            {
                try
                {
                    LogAppendValue(Color.White, Color.DarkBlue, "Dodawanie atrybutów");
                    XLAtrybutInfo_20231 atrybut1 = new XLAtrybutInfo_20231();
                    atrybut1.Wersja = XLWersjaAPI;
                    atrybut1.GIDFirma = xlDokumentZamInfo.GIDFirma;
                    atrybut1.GIDLp = xlDokumentZamInfo.GIDLp;
                    atrybut1.GIDNumer = xlDokumentZamInfo.GIDNumer;
                    atrybut1.GIDSubLp = 0;
                    atrybut1.GIDTyp = xlDokumentZamInfo.GIDTyp;
                    atrybut1.Wartosc = dostawa;
                    atrybut1.Klasa = "Nr dostawy";
                    int blad_atr = cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut1);

                    if (blad_atr != 0)
                    {
                        MessageBox.Show("Błąd atrybutu nr dostawy: " + blad_atr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    XLAtrybutInfo_20231 atrybut2 = new XLAtrybutInfo_20231();
                    atrybut2.Wersja = XLWersjaAPI;
                    atrybut2.GIDFirma = xlDokumentZamInfo.GIDFirma;
                    atrybut2.GIDLp = xlDokumentZamInfo.GIDLp;
                    atrybut2.GIDNumer = xlDokumentZamInfo.GIDNumer;
                    atrybut2.GIDSubLp = 0;
                    atrybut2.GIDTyp = xlDokumentZamInfo.GIDTyp;
                    atrybut2.Wartosc = "0";
                    atrybut2.Klasa = "Zużyta ilość";
                    blad_atr = cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut2);

                    if (blad_atr != 0)
                    {
                        MessageBox.Show("Błąd atrybutu nr dostawy: " + blad_atr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                   
                }
                catch (Exception e)
                {
                    MessageBox.Show("Błąd dodawania atrybutów do ZS: " + e.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }



            LogAppendValue(Color.Black, Color.White,$@"Dodano pozycje o nazwie {xlDokumentZamInfo.Nazwa} i kodzie {str} w ilosci {ilosc}");
            if (str != 0)
            {
                XLKomunikatInfo_20231 kom = new XLKomunikatInfo_20231();
                kom.Funkcja = 42;
                kom.Wersja = XLWersjaAPI;
                kom.Blad = str;
                cdn_api.cdn_api.XLOpisBledu(kom);
                LogAppendValue(Color.DarkCyan, Color.White, $@"Kod API: {str} Opis błędu {kom.OpisBledu}");

            }
            return str;
        }

        public int AddElement_XL_ZZ(int IDDokumentuZam, CSVList csvlist, string ilosc, string grupaPod, string typDok)
        {
            int str=0;
            int pozostalaIlosc = Convert.ToInt32(ilosc);
            while (pozostalaIlosc > 0)
            {
                int[] index = AddIndex_XL_ZZ(csvlist, pickedDate, RodzajDzialania, grupaPod);
                //Twr_GIDNumer , Twr_GIDTyp   , Twr_GIDLP  , Twr_GIDFirma ,ZaE_GIDNumer  , ZaE_GIDTyp   , ZaE_GIDLp  ,ZaE_GIDFirma, strybut IloscUzyta, ZaE_Ilosc 
                
                    int iloscPoz = index[9] - index[8];
                    if (iloscPoz > pozostalaIlosc || iloscPoz==0)
                        iloscPoz = pozostalaIlosc;
                
               

                XLDokumentZamElemInfo_20231 xlDokumentZamInfo = new XLDokumentZamElemInfo_20231
                {
                    Wersja = XLWersjaAPI,
                    Towar = csvlist.Articlegroup,
                    Nazwa = csvlist.Article,
                    TowarEAN = csvlist.Barcode,
                    Ilosc = Convert.ToString(iloscPoz),
                    TwrNumer = index[0],
                    TwrTyp = index[1],
                    TwrLp = index[2],
                    TwrFirma = index[3],
                    CenaOferowana = Regex.Replace(csvlist.ClientVal, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", "."),
                    CenaKatalogowa = Regex.Replace(csvlist.Price, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", "."),
                    Rabat="0",
                    RabatReq=1
                };

                if (grupaPod == "2")
                {
                    xlDokumentZamInfo.GrupaVAT = "C";
                    xlDokumentZamInfo.StawkaVAT = "0";
                }

                try
                {
                    str = cdn_api.cdn_api.XLDodajPozycjeZam(IDDokumentuZam, xlDokumentZamInfo);
                }
                catch (Exception e)
                {
                    throw new Exception("Błąd nr " + e.ToString() + " kod API: " + str);
                }

                pozostalaIlosc = pozostalaIlosc - iloscPoz;
                if (index[7] != 0)
                {

                    string updateIlosc = "Update a set Atr_Wartosc = '"+Convert.ToString(iloscPoz+index[8])+"' FROM cdn.atrybuty as a join cdn.ZamElem as e on ZaE_GIDFirma = a.Atr_ObiFirma and ZaE_GIDLp = a.Atr_ObiLp and ZaE_GIDNumer = a.Atr_ObiNumer and ZaE_GIDTyp = a.Atr_ObiTyp "
                         + "WHERE Atr_AtkId = 51 and ZaE_GIDFirma = "+index[7]+" and ZaE_GIDLp = "+index[6]+" and ZaE_GIDNumer = "+index[4]+" and ZaE_GIDTyp = "+index[5];
                    Update(updateIlosc);
                    
                }

                LogAppendValue(Color.Black, Color.White, $@"Dodano pozycje o nazwie {xlDokumentZamInfo.Nazwa} i kodzie {str} w ilosci {ilosc}");
                if (str != 0)
                {
                    XLKomunikatInfo_20231 kom = new XLKomunikatInfo_20231();
                    kom.Funkcja = 42;
                    kom.Wersja = XLWersjaAPI;
                    kom.Blad = str;
                    cdn_api.cdn_api.XLOpisBledu(kom);
                    LogAppendValue(Color.DarkCyan, Color.White, $@"Kod API: {str} Opis błędu {kom.OpisBledu}");

                }
            }
            return str;
        }
        //dodawanie towaru z zamowienia do katalogu

        public int[] AddIndex_XL(CSVList order, DateTime pickedTime,int tryb, string grupaPodatkowa)
        {
            AttachThreadToClarion(true);
             
            int return_value = -1;
            bool is_tf = false;
            string qry_resp;

            if (ConnectToSQL(connection_string, $"begin if exists(Select 1 from cdn.TwrKarty where Twr_Ean='{order.Barcode}' and Twr_Archiwalny=1) select top 1 1 else select top 1 0 end") != "0")
            {
                string updateEAN = $"UPDATE cdn.TwrKarty set Twr_EAN = '' WHERE Twr_Ean = '{order.Barcode}' and Twr_Archiwalny = 1";
                Update(updateEAN);
                LogAppendValue(Color.Black, Color.White, $@"Usunięto EAN {order.Barcode} o gid dla towarów archiwalnych");
            }

            if (ConnectToSQL(connection_string, $"begin if exists(Select 1 from cdn.TwrKarty where Twr_Ean='{order.Barcode}') select top 1 1 else select top 1 0 end") != "0")
            {
                try
                {
                    int[] knt_arr = ConnectListToSQL(connection_string, $@"SELECT TOP 1 Cast(Twr_GIDNumer as varchar) 
                      , Cast(Twr_GIDTyp as varchar) 
                      , Cast(Twr_GIDLP as varchar) 
                      , Cast(Twr_GIDFirma as varchar) 
                    FROM cdn.TwrKarty where Twr_Ean = '{order.Barcode}'").ToArray();
                    LogAppendValue(Color.Black, Color.White, $@"Odnaleziono towar dla EAN {order.Barcode} o gid {knt_arr[0]}");
                    if (RodzajDzialania == 1)
                    {
                        string updateQuery = "update cdn.TwrKarty set Twr_EAN='' where Twr_GIDNumer=" + knt_arr[0] + " and Twr_GIDTyp=" + knt_arr[1] + " and Twr_GIDLp= " + knt_arr[2] + " and Twr_GIDFirma=" + knt_arr[3];
                        Update(updateQuery);
                        LogAppendValue(Color.Black, Color.White, $@"Usunięto EAN {order.Barcode} o gid {knt_arr[0]}");
                    }
                    else
                        return knt_arr;
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed,$@"Wystąpił błąd rzutowania kontrahenta {e}");
                }
            }
            string ogrodnik = "";
            if (order.Grower.Replace($@"'", $@"''").Length>1)
            {


            ogrodnik = ConnectToSQL(connection_string, $@"declare @resp as VARCHAR(MAX) 
            declare @grower as VARCHAR(MAX) 
            set @grower='{order.Grower.Replace($@"'", $@"''")}' 
            if not exists (select 1 from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower) 
                        begin 
			            insert into cdn.AtrybutyWartosci(AtW_AtKId, AtW_Wartosc, AtW_OddZrdId, AtW_Lp) 
                        Values( (select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2'), @grower, 0, 0); 
			            set @resp = Right('00000' + Convert(varchar(5),(select top 1 AtW_ID from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower)),5) 
			            end 
			            else 
			            begin 
                        set @resp = Right('00000' + Convert(varchar(5),(select top 1 AtW_ID from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower)),5) 
			            end select @resp");
            if (ogrodnik.Length > 10)
            {
                LogAppendValue(Color.White, Color.DarkRed,$@"Wystąpił błąd przy dodawaniu id ogrodnika do atrybutu Hodowca2 {ogrodnik} !");
                    ogrodnik = "";
            }
            }
            string nazwa = $@"{order.Articleno}-{order.Mark1}-{order.Mark2}-{order.Mark3}_{pickedTime.ToString("yyMMdd")}_{ogrodnik}";
            nazwa = Truncate(nazwa, 35);
            
            if (order.Subdeb.Length == 0 || tryb == 2)
            {
                nazwa += "_TUFL";
                is_tf = true;
            }
            else
                nazwa += ("_" + order.Subdeb);
            //Sprawdzic czy istnieje duplikat
            qry_resp = ConnectToSQL(connection_string, $@"declare @reszta_kodu varchar(max) 
                declare @vbn varchar(max) 
                declare @srodek varchar(max) 
                declare @count int 
                set @vbn = '{nazwa.Substring(0, nazwa.IndexOf('-'))}' 
                set @srodek = '' 
                set @reszta_kodu ='{nazwa.Substring(nazwa.IndexOf('-'))}' 
                set @count = 0 
                if not exists(Select 1 from cdn.TwrKarty where Twr_Kod=@vbn+@srodek+@reszta_kodu) 
	                begin 
	                select 0 
	                end 
                else 
	                begin 
	                while (Select 1 from cdn.TwrKarty where Twr_Kod=@vbn+@srodek+@reszta_kodu)=1  
	                begin 
	                set @count = @count+1 
	                set @srodek = '|'+CAST(@count as varchar(max)) 
	                end 
                select @count 
                end");
            if (qry_resp != "0" && int.TryParse(qry_resp, out int number) == true)
            {
                LogAppendValue(Color.Black, Color.White, $@"Znaleziono duplikat nr {number} towaru {nazwa} o EAN {order.Barcode}");
                try
                {
                    nazwa = $@"{order.Articleno}|{number}-{order.Mark1}-{order.Mark2}-{order.Mark3}_{pickedTime.ToString("yyMMdd")}_{ogrodnik}";
                    nazwa = Truncate(nazwa, 35);
                    if (order.Subdeb.Length == 0 || tryb == 2)
                    {
                        nazwa += "_TUFL";
                    }
                    else
                        nazwa += ("_" + order.Subdeb);
                }
                catch (Exception ex)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd {++errorCount} przy duplikacie {nazwa} z EAN {order.Barcode} {ex}!");
                    return new int[4] { 0, 0, 0, 0 };
                }
            }
            else if (int.TryParse(qry_resp, out int numberek) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd {++errorCount} rzutowania dla Query {qry_resp}");
            //Stawka podatkowa z kodu Code - jest tabelka w XL
            XLTowarInfo_20231 xLTowarInfo = new XLTowarInfo_20231
            {
                Wersja = XLWersjaAPI,
                Typ = 1,
                Kod = nazwa,
                StawkaPod = "23",
                GrupaPod = "A",
                StawkaPodSpr = "23",
                GrupaPodSpr = "A",
                Nazwa = order.Article,
                EAN = order.Barcode,
                PCN = order.Code.PadLeft(8,'0'),
                SWW = Regex.Replace(order.Code, ".{2}", "$0.").TrimEnd('.'),
                //TwrGrupa = "KWIATY DEDYKOWANE",
                Jm = "szt.", 
            };
            LogAppendValue(Color.Black, Color.White, $@"Dodawanie nowego towaru o kodzie {nazwa} oraz nazwie {order.Article}");
            try
            {
                string stawka = "23";
                var resp = ConnectListStrToSQL(connection_string, $@"IF EXISTS (select 1 from cdn.KodyCN where KCN_Kod='{order.Code.PadLeft(8, '0')}') 
                begin 
                select top 1 Cast(KCN_Kod as varchar), 
			                 Cast(KCN_StawkaVAT as varchar), 
			                 Cast(KCN_Opis as varchar(max)) 
			                 from cdn.KodyCN where KCN_Kod='{order.Code.PadLeft(8, '0')}' 
                end 
                else 
                select 'x', 'x', 'x'").ToArray();
                //MessageBox.Show($@"{resp[0]}, {resp[1]}, {resp[2]}");
                LogAppendValue(Color.Black, Color.White,$@"Dodawanie kodu KCN, stawka i opisu: {resp[0]}, {resp[1]}, {resp[2].Split('#').LastOrDefault()}");
                if (resp[0].Length > 1 && resp[1].Length>0)
                {
                    xLTowarInfo.PCN = resp[0];
                    stawka = Regex.Replace(resp[1], @"[^\d]", "").TrimEnd('0');
                    switch (stawka)
                    {
                        case "23":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "A";
                            xLTowarInfo.GrupaPodSpr = "A";
                            break;
                        case "8":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "B";
                            xLTowarInfo.GrupaPodSpr = "B";
                            break;
                        case "5":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "G";
                            xLTowarInfo.GrupaPodSpr = "G";
                            break;
                        case "3":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "F";
                            xLTowarInfo.GrupaPodSpr = "F";
                            break;
                        case "0":

                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "C";
                            xLTowarInfo.GrupaPodSpr = "C";
                            break;
                        default:
                            xLTowarInfo.StawkaPod = "23";
                            xLTowarInfo.StawkaPodSpr = "23";
                            xLTowarInfo.GrupaPod = "A";
                            xLTowarInfo.GrupaPodSpr = "A";
                            break;
                    }
                    xLTowarInfo.SWW = resp[2].Split('#').LastOrDefault();
                }
            }
            catch (Exception exce)
            {
                LogAppendValue(Color.White, Color.DarkRed,$@"Wystąpił błąd nr {++errorCount} konwersji podatkowej {exce} " + "\n");
            }
            if (grupaPodatkowa == "2")
            {
                xLTowarInfo.StawkaPod = "0";
                xLTowarInfo.GrupaPod = "C";
            }
            if (is_tf)
            {
                try
                {
                    string group = FindGroup(order.Article.Replace($@"'", $@"''"));
                    LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $@"Odnaleziona nazwa grupy to {group}");
                    var resp = ConnectToSQL(GetConnectionString(), $@"IF EXISTS (select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa='{group}' and [CDN].[TwrGrupaPelnaNazwa] (TwG_GIDNumer) like '%Kwiaty Cięte/%') 
                    begin  
                    select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa='{group}' and [CDN].[TwrGrupaPelnaNazwa] (TwG_GIDNumer) like '%Kwiaty Cięte/%'
                    end 
                    else 
                    begin 
                    select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Kod='{group}' and [CDN].[TwrGrupaPelnaNazwa] (TwG_GIDNumer) like '%Kwiaty Cięte/%' 
                    end");
                    LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $"Odnalezione GID to {resp}");

                    string[] completeGrupa = resp.Split(';');
                    if (completeGrupa.Length > 0 && group != "")
                    {
                        xLTowarInfo.TwGNumer = Int32.Parse(completeGrupa[0]);
                        xLTowarInfo.TwGTyp = Int32.Parse(completeGrupa[1]);
                        xLTowarInfo.TwGLp = Int32.Parse(completeGrupa[2]);
                        xLTowarInfo.TwGFirma = Int32.Parse(completeGrupa[3]);
                    }
                    else
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Błąd Nie znaleziono czegos dla towaru {xLTowarInfo.Nazwa}, odpowiedz to {resp}");
                        xLTowarInfo.TwGNumer = 26;
                        xLTowarInfo.TwGTyp = -16;
                        xLTowarInfo.TwGLp = 0;
                        xLTowarInfo.TwGFirma = 1403393;
                    }
 
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} dodawania grupy towarowej {e} " + "\n");
                }
            }
            else xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE";
            try
            {
                return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);

            }
            catch (Exception e)
            {
                throw new Exception("Wystapil wyjatek przy dodawaniu towaru. Kod API: " + return_value + ", Wyjątek: " + e);

            }
            if (return_value != 0)
            {
                if (return_value != 1) LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} dodania towaru {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");
                LogAppendValue(Color.Black, Color.White, $"Poprwka wprowadzania grupy dla towaru {nazwa}");
                xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE";
                return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
                string group = ConnectToSQL(connection_string, "select top 1 CONCAT(TwG_GIDNumer, ';', TwG_GIDTyp, ';', TwG_GIDLp, ';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Kod = 'KWIATY DEDYKOWANE'");
                if (return_value != 0)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} poprawki x2 towaru <Z nazwa> {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                try
                    {
                        int[] spltGrp = group.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                           .Select(x => int.Parse(x))
                           .ToArray();
                        if (spltGrp.Length >= 3)
                        {
                            xLTowarInfo.TwGNumer = spltGrp[0];
                            xLTowarInfo.TwGTyp = spltGrp[1];
                            xLTowarInfo.TwGLp = spltGrp[2];
                            xLTowarInfo.TwGFirma = spltGrp[3];
                        }
                        return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
                        if (return_value != 0) LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} poprawki x3 towaru <z GID> {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                    }
                    catch (Exception)
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} pobierania GID grupy KWIATY DEDYKOWANE: {xLTowarInfo.Kod} z TwGNumer {xLTowarInfo.GIDNumer} TwGNumer {xLTowarInfo.TwGNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                    }
                }


            }

            XLZamkniecieTowaruInfo_20231 xLZamkniecieTowaruInfo_20231 = new XLZamkniecieTowaruInfo_20231
            {
                Wersja = XLWersjaAPI,
                Tryb = return_value != 0 ? -1 : 0,
            };
            return_value = cdn_api.cdn_api.XLZamknijTowar(XLtowar_id, xLZamkniecieTowaruInfo_20231);
            Update("UPDATE  c SET TwC_TcnId = 2, [TwC_DataOd]=986562792 FROM [CDN].[TwrCeny] as c join cdn.TwrCenyNag on TwC_TcnId=TCN_Id WHERE TwC_TwrTyp = 16 AND TwC_TwrFirma = " + xLTowarInfo.GIDFirma + " AND TwC_TwrNumer = " + xLTowarInfo.GIDNumer + " and TCN_RodzajCeny=2");
            Update("UPDATE  c SET TwC_TcnId = 1, [TwC_DataOd]=986562792 FROM [CDN].[TwrCeny] as c join cdn.TwrCenyNag on TwC_TcnId=TCN_Id WHERE TwC_TwrTyp = 16 AND TwC_TwrFirma = " + xLTowarInfo.GIDFirma + " AND TwC_TwrNumer = " + xLTowarInfo.GIDNumer + " and TCN_RodzajCeny=1");


            if (return_value != 0)
            {
                LogAppendValue(Color.White, Color.DarkRed, $"Błąd przy zamknięciu towaru towar : {xLTowarInfo.Kod} dodany do magazynu o gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");

            }
            else
            {
                LogAppendValue(Color.Black, Color.White,$@"Utworzono towar z nazwą {xLTowarInfo.Nazwa} o kodzie {xLTowarInfo.Kod}");
                if (is_tf == true)
                {
                    string resp = ConnectToSQL(connection_string, $@"update t 
                    set  Twr_ESklep = 1 
                    from cdn.TwrKarty as t 
                    where t.Twr_GIDNumer = {xLTowarInfo.GIDNumer} and t.Twr_GIDTyp={xLTowarInfo.GIDTyp} and t.Twr_GIDFirma={xLTowarInfo.GIDFirma}");
                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do e-Sklepu {"\n"} {resp}");
                    }

                    resp = ConnectToSQL(connection_string, $@"update cdn.TwrAplikacje set TAP_Status=1, TAP_Dostepnosc=1, [TAP_Flagi] = 66 ,
                    TAP_GrupaeSklep=(select top 1 TwG_GrONumer from cdn.TwrGrupy where TwG_GIDNumer={xLTowarInfo.GIDNumer} and TwG_GIDTyp ={xLTowarInfo.GIDTyp} order by TwG_SyncId desc)  where TAP_ObiNumer = {xLTowarInfo.GIDNumer} and TAP_ObiTyp={xLTowarInfo.GIDTyp}");
                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do TwrAplikacje {"\n"} {resp}");
                    }
                    

                }
               /* else
                {
                    string resp = ConnectToSQL(connection_string, $@"update t 
                    set Twr_MobSpr = 1  
                    from cdn.TwrKarty as t 
                    where t.Twr_GIDNumer = {xLTowarInfo.GIDNumer} and t.Twr_GIDTyp={xLTowarInfo.GIDTyp} and t.Twr_GIDFirma={xLTowarInfo.GIDFirma}");

                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do e-Sklepu {"\n"} {resp}");
                    }
                }*/

                    string msg = ConnectToSQL(connection_string, "update cdn.TwrCeny set TwC_Wartosc=Cast('" + Regex.Replace(order.Price, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", ".") + "' as decimal(18,4)), TwC_Marza='0.00' where TwC_TwrNumer = (select Twr_GIDNumer from cdn.TwrKarty where Twr_Kod = '" + xLTowarInfo.Kod + "') and TwC_TwrLp=1");
                    if (msg != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $"błąd przy aktualizacji ceny {msg}!");
                    }
                    msg = ConnectToSQL(connection_string, $@"update cdn.TwrKarty set TwR_EAN='{order.Barcode}' , Twr_MobSpr = 1 where Twr_GIDNumer ={xLTowarInfo.GIDNumer} and Twr_GIDtyp={xLTowarInfo.GIDTyp}");
                    if (msg != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $"Błąd przy aktualizacji kodu EAN (wiodące zera z przodu) {msg}!");
                    } 

                    if (return_value == 0)
                    {
                    try
                    {
                        //{ "VBN", "Klient", "Hodowca2", "S1", "S2", "S3", "IMG_URL" }
                        //Dodawanie Atrybutów
                        string flowerType = ConnectToSQL(connection_string, $"select top 1 Typ from FlowerType where KodCN={order.Code.PadLeft(8, '0')}");
                        LogAppendValue(Color.Black, Color.White, $"Rodzaj towaru dla Kodu CN {order.Code.PadLeft(8, '0')} to  {flowerType}");
                        var atr_resp = AddAtribute(new string[8] { order.Articleno, order.Grower, order.Clientnumber, order.Mark1, order.Mark2, "", order.Mark3, order.Picture }, new int[4] { xLTowarInfo.GIDFirma, xLTowarInfo.GIDLp, xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp }, flowerType);
                        if (atr_resp.Length > 0)
                        {
                            LogAppendValue(Color.White, Color.DarkBlue, $"Bład SQL przy dodawaniu atrybutów {atr_resp}");
                        }
                        else
                        {
                            LogAppendValue(Color.Black, Color.White, $"Pomyślnie dodano atrybuty");
                        }

                    }
                    catch (Exception e)
                    {
                        LogAppendValue(Color.Black, Color.White, "Nie udało się załadować obrazu " + e);
                    }
                    try
                    {
                        if (order.Picture.Length > 0)
                        {
                           // LogAppendValue(Color.DeepPink, Color.White, "Dodawanie Picture");
                            var img = DownloadImage(order.Picture);
                            if (img != null)
                            {
                                UploadImage(xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp, ((img.Size.Width > 512 && img.Size.Height > 512) ? ImageToByte2(ResizeImage(img, 512, 512)) : ImageToByte2(img)), order.Article, order.Articleno);
                            }
                            //Zmienianie Atrybutów
                            string atr_Resp = ConnectToSQL(connection_string,
                                $@"update cdn.DaneObiekty set DAO_MobSpr = 1 
                                    where DAO_ObiNumer = (select Twr_GIDNumer from cdn.TwrKarty where Twr_Kod = '{xLTowarInfo.Kod}')");
                            if (atr_Resp != "")
                            {
                                LogAppendValue(Color.White, Color.DarkBlue, "Nie udało się dodać atrybutu DAO_mobspr, odp. SQL to " + atr_Resp);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $"Błąd nr {++errorCount}Nie udało się dodać obrazu " + ex);
                    }
                }

                if (ZZ_knt[0]>0)
                {
                string a = ConnectToSQL(connection_string, $@"declare @nip varchar(10)
                    set @nip = (select top 1 Knt_NipPrefiks from cdn.KntKarty where Knt_GIDNumer={ZZ_knt[0]} order by Knt_GIDNumer desc)
                    if @nip is not null 
                    begin 
                    update cdn.TwrKarty set Twr_KrajPoch=@nip where Twr_GIDNumer={xLTowarInfo.GIDNumer} and Twr_GIDLp={xLTowarInfo.GIDLp} and Twr_GIDTyp={xLTowarInfo.GIDTyp} and Twr_GIDFirma={xLTowarInfo.GIDFirma} end");
                if (a != "") LogAppendValue(Color.White, Color.DarkRed,$@"Wystąpił błąd nr {++errorCount} dodawania kraju {a}");
                }
            }
            return new int[4] { xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp, xLTowarInfo.GIDLp, xLTowarInfo.GIDFirma };
        }



        public int[] AddIndex_XL_ZZ(CSVList order, DateTime pickedTime, int tryb, string grupaPodatkowa)
        {
            AttachThreadToClarion(true);

            int return_value = -1;
            bool is_tf = false;
            string qry_resp;
            string czyIstnieje = ConnectToSQL(connection_string, "begin if exists(SELECT TOP (1) 1  FROM cdn.ZamElem as e "
                   + "join cdn.atrybuty as a1 on ZaE_GIDFirma = a1.Atr_ObiFirma and ZaE_GIDLp = a1.Atr_ObiLp and ZaE_GIDNumer = Atr_ObiNumer and ZaE_GIDTyp = a1.Atr_ObiTyp "
                   + "join cdn.atrybuty as a3 on ZaE_GIDFirma = a3.Atr_ObiFirma and ZaE_GIDLp = a3.Atr_ObiLp and ZaE_GIDNumer = a3.Atr_ObiNumer and ZaE_GIDTyp = a3.Atr_ObiTyp "
                   + "join cdn.TwrKarty as t on ZaE_TwrFirma = Twr_GIDFirma and e.ZaE_TwrLp = t.Twr_GIDLp and e.ZaE_TwrNumer = t.Twr_GIDNumer and e.ZaE_TwrTyp = t.Twr_GIDTyp "
                   + "join cdn.atrybuty as a2 on Twr_GIDFirma = a2.Atr_ObiFirma and Twr_GIDLp = a2.Atr_ObiLp and Twr_GIDNumer = a2.Atr_ObiNumer and Twr_GIDTyp = a2.Atr_ObiTyp "
                   + "where a1.Atr_AtkId = 50 and a1.Atr_Wartosc = '" + dostawa + "' and a3.Atr_AtkId = 51 and a3.Atr_Wartosc < e.ZaE_Ilosc and a2.Atr_AtkId = 52 and a2.Atr_Wartosc = '" + order.Barcode + "') select 1 else select 0 end");
            LogAppendValue(Color.Black, Color.White, $@"CzyIstnieje {czyIstnieje}");
            if (czyIstnieje.Contains("1"))
            {
                try
                {
                    int[] knt_arr = ConnectListToSQLZZ(connection_string, $@"SELECT TOP(1) Cast(Twr_GIDNumer as varchar)  , Cast(Twr_GIDTyp as varchar)  , Cast(Twr_GIDLP as varchar)  , Cast(Twr_GIDFirma as varchar) ,
	                                Cast(ZaE_GIDNumer as varchar)  , Cast(ZaE_GIDTyp as varchar)  , Cast(ZaE_GIDLp as varchar)  , Cast(ZaE_GIDFirma as varchar) ,
	                                Cast(a3.atr_Wartosc as varchar), SUBSTRING(Cast(ZaE_Ilosc as varchar), 0, CHARINDEX('.', Cast(ZaE_Ilosc as varchar)))
                                    FROM cdn.ZamElem as e
                                    join cdn.atrybuty as a1 on ZaE_GIDFirma=a1.Atr_ObiFirma and ZaE_GIDLp=a1.Atr_ObiLp and ZaE_GIDNumer=Atr_ObiNumer and ZaE_GIDTyp=a1.Atr_ObiTyp
                                    join cdn.atrybuty as a3 on ZaE_GIDFirma=a3.Atr_ObiFirma and ZaE_GIDLp=a3.Atr_ObiLp and ZaE_GIDNumer=a3.Atr_ObiNumer and ZaE_GIDTyp=a3.Atr_ObiTyp
                                    join cdn.TwrKarty as t on ZaE_TwrFirma=Twr_GIDFirma and e.ZaE_TwrLp=t.Twr_GIDLp and e.ZaE_TwrNumer=t.Twr_GIDNumer and e.ZaE_TwrTyp=t.Twr_GIDTyp
                                    join cdn.atrybuty as a2 on Twr_GIDFirma=a2.Atr_ObiFirma and Twr_GIDLp=a2.Atr_ObiLp and Twr_GIDNumer=a2.Atr_ObiNumer and Twr_GIDTyp=a2.Atr_ObiTyp
                                    where a1.Atr_AtkId=50 and a1.Atr_Wartosc='{dostawa}' and a3.Atr_AtkId=51 and a3.Atr_Wartosc<e.ZaE_Ilosc and a2.Atr_AtkId=52 and a2.Atr_Wartosc='{order.Barcode}'
                                    ORDER BY ZaE_GIDNumer").ToArray();
                    LogAppendValue(Color.Black, Color.White, $@"Odnaleziono towar dla EAN {order.Barcode} o gid {knt_arr[0]}");
                   
                        return knt_arr;
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd rzutowania kontrahenta {e}");
                }
            }


            if (ConnectToSQL(connection_string, $"begin if exists(Select 1 from cdn.TwrKarty where Twr_Ean='{order.Barcode}' and Twr_Archiwalny=1) select top 1 1 else select top 1 0 end") != "0")
            {
                string updateEAN = $"UPDATE cdn.TwrKarty set Twr_EAN = '' WHERE Twr_Ean = '{order.Barcode}' and Twr_Archiwalny = 1";
                Update(updateEAN);
                LogAppendValue(Color.Black, Color.White, $@"Usunięto EAN {order.Barcode} o gid dla towarów archiwalnych");
            }

            if (ConnectToSQL(connection_string, "begin if exists(Select 1 from cdn.TwrKarty where Twr_Ean='{order.Barcode}') select top 1 '1' else select top 1 '0' end") != "0")
            {
                try
                {
                    int[] knt_arr = ConnectListToSQL(connection_string, $@"SELECT TOP 1 Cast(Twr_GIDNumer as varchar) 
                      , Cast(Twr_GIDTyp as varchar) 
                      , Cast(Twr_GIDLP as varchar) 
                      , Cast(Twr_GIDFirma as varchar), '0', '0', '0', '0', '0', '0'
                    FROM cdn.TwrKarty where Twr_Ean = '{order.Barcode}'").ToArray();
                    LogAppendValue(Color.Black, Color.White, $@"Odnaleziono towar dla EAN {order.Barcode} o gid {knt_arr[0]}");

                    if (ConnectToSQL(connection_string, "SELECT Twr_Kod From cdn.TwrKarty WHERE Twr_GIDNumer=" + knt_arr[0] + " and Twr_GIDTyp=" + knt_arr[1] + " and Twr_GIDLp= " + knt_arr[2] + " and Twr_GIDFirma=" + knt_arr[3]).Contains("TUFL"))
                        return knt_arr;
                    else
                    {
                        string updateQuery = "update cdn.TwrKarty set Twr_EAN='' where Twr_GIDNumer=" + knt_arr[0] + " and Twr_GIDTyp=" + knt_arr[1] + " and Twr_GIDLp= " + knt_arr[2] + " and Twr_GIDFirma=" + knt_arr[3];
                        Update(updateQuery);
                        LogAppendValue(Color.Black, Color.White, $@"Usunięto EAN {order.Barcode} o gid {knt_arr[0]}");
                    }
                    
                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd usuwania EAN {e}");
                    return new int[10] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                }
            }
            string ogrodnik = "";
            if (order.Grower.Replace($@"'", $@"''").Length > 1)
            {


                ogrodnik = ConnectToSQL(connection_string, $@"declare @resp as VARCHAR(MAX) 
            declare @grower as VARCHAR(MAX) 
            set @grower='{order.Grower.Replace($@"'", $@"''")}' 
            if not exists (select 1 from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower) 
                        begin 
			            insert into cdn.AtrybutyWartosci(AtW_AtKId, AtW_Wartosc, AtW_OddZrdId, AtW_Lp) 
                        Values( (select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2'), @grower, 0, 0); 
			            set @resp = Right('00000' + Convert(varchar(5),(select top 1 AtW_ID from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower)),5) 
			            end 
			            else 
			            begin 
                        set @resp = Right('00000' + Convert(varchar(5),(select top 1 AtW_ID from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower)),5) 
			            end select @resp");
                if (ogrodnik.Length > 10)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd przy dodawaniu id ogrodnika do atrybutu Hodowca2 {ogrodnik} !");
                    ogrodnik = "";
                }
            }
            string nazwa = $@"{order.Articleno}-{order.Mark1}-{order.Mark2}-{order.Mark3}_{pickedTime.ToString("yyMMdd")}_{ogrodnik}";
            nazwa = Truncate(nazwa, 35);

            if (order.Subdeb.Length == 0 || tryb == 2)
            {
                nazwa += "_TUFL";
                is_tf = true;
            }
            else
                nazwa += ("_" + order.Subdeb);
            //Sprawdzic czy istnieje duplikat
            qry_resp = ConnectToSQL(connection_string, $@"declare @reszta_kodu varchar(max) 
                declare @vbn varchar(max) 
                declare @srodek varchar(max) 
                declare @count int 
                set @vbn = '{nazwa.Substring(0, nazwa.IndexOf('-'))}' 
                set @srodek = '' 
                set @reszta_kodu ='{nazwa.Substring(nazwa.IndexOf('-'))}' 
                set @count = 0 
                if not exists(Select 1 from cdn.TwrKarty where Twr_Kod=@vbn+@srodek+@reszta_kodu) 
	                begin 
	                select 0 
	                end 
                else 
	                begin 
	                while (Select 1 from cdn.TwrKarty where Twr_Kod=@vbn+@srodek+@reszta_kodu)=1  
	                begin 
	                set @count = @count+1 
	                set @srodek = '|'+CAST(@count as varchar(max)) 
	                end 
                select @count 
                end");
            if (qry_resp != "0" && int.TryParse(qry_resp, out int number) == true)
            {
                LogAppendValue(Color.Black, Color.White, $@"Znaleziono duplikat nr {number} towaru {nazwa} o EAN {order.Barcode}");
                try
                {
                    nazwa = $@"{order.Articleno}|{number}-{order.Mark1}-{order.Mark2}-{order.Mark3}_{pickedTime.ToString("yyMMdd")}_{ogrodnik}";
                    nazwa = Truncate(nazwa, 35);
                    if (order.Subdeb.Length == 0 || tryb == 2)
                    {
                        nazwa += "_TUFL";
                    }
                    else
                        nazwa += ("_" + order.Subdeb);
                }
                catch (Exception ex)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd {++errorCount} przy duplikacie {nazwa} z EAN {order.Barcode} {ex}!");
                    return new int[10] { 0, 0, 0, 0 , 0, 0, 0, 0, 0, 0};
                }
            }
            else if (int.TryParse(qry_resp, out int numberek) == false) LogAppendValue(Color.White, Color.DarkRed, $@"Błąd {++errorCount} rzutowania dla Query {qry_resp}");
            //Stawka podatkowa z kodu Code - jest tabelka w XL
            XLTowarInfo_20231 xLTowarInfo = new XLTowarInfo_20231
            {
                Wersja = XLWersjaAPI,
                Typ = 1,
                Kod = nazwa,
                StawkaPod = "23",
                GrupaPod = "A",
                StawkaPodSpr = "23",
                GrupaPodSpr = "A",
                Nazwa = order.Article,
                EAN = order.Barcode,
                PCN = order.Code.PadLeft(8, '0'),
                SWW = Regex.Replace(order.Code, ".{2}", "$0.").TrimEnd('.'),
                //TwrGrupa = "KWIATY DEDYKOWANE",
                Jm = "szt.",
            };
            LogAppendValue(Color.Black, Color.White, $@"Dodawanie nowego towaru o kodzie {nazwa} oraz nazwie {order.Article}");
            try
            {
                string stawka = "23";
                var resp = ConnectListStrToSQL(connection_string, $@"IF EXISTS (select 1 from cdn.KodyCN where KCN_Kod='{order.Code.PadLeft(8, '0')}') 
                begin 
                select top 1 Cast(KCN_Kod as varchar), 
			                 Cast(KCN_StawkaVAT as varchar), 
			                 Cast(KCN_Opis as varchar(max)) 
			                 from cdn.KodyCN where KCN_Kod='{order.Code.PadLeft(8, '0')}' 
                end 
                else 
                select 'x', 'x', 'x'").ToArray();
                //MessageBox.Show($@"{resp[0]}, {resp[1]}, {resp[2]}");
                LogAppendValue(Color.Black, Color.White, $@"Dodawanie kodu KCN, stawka i opisu: {resp[0]}, {resp[1]}, {resp[2].Split('#').LastOrDefault()}");
                if (resp[0].Length > 1 && resp[1].Length > 0)
                {
                    xLTowarInfo.PCN = resp[0];
                    stawka = Regex.Replace(resp[1], @"[^\d]", "").TrimEnd('0');
                    switch (stawka)
                    {
                        case "23":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "A";
                            xLTowarInfo.GrupaPodSpr = "A";
                            break;
                        case "8":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "B";
                            xLTowarInfo.GrupaPodSpr = "B";
                            break;
                        case "5":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "G";
                            xLTowarInfo.GrupaPodSpr = "G";
                            break;
                        case "3":
                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "F";
                            xLTowarInfo.GrupaPodSpr = "F";
                            break;
                        case "0":

                            xLTowarInfo.StawkaPod = stawka;
                            xLTowarInfo.StawkaPodSpr = stawka;
                            xLTowarInfo.GrupaPod = "C";
                            xLTowarInfo.GrupaPodSpr = "C";
                            break;
                        default:
                            xLTowarInfo.StawkaPod = "23";
                            xLTowarInfo.StawkaPodSpr = "23";
                            xLTowarInfo.GrupaPod = "A";
                            xLTowarInfo.GrupaPodSpr = "A";
                            break;
                    }
                    xLTowarInfo.SWW = resp[2].Split('#').LastOrDefault();
                }
            }
            catch (Exception exce)
            {
                LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} konwersji podatkowej {exce} " + "\n");
            }
            if (grupaPodatkowa == "2")
            {
                xLTowarInfo.StawkaPod = "0";
                xLTowarInfo.GrupaPod = "C";
            }
            if (is_tf)
            {
                try
                {
                    string group = FindGroup(order.Article.Replace($@"'", $@"''"));
                    LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $@"Odnaleziona nazwa grupy to {group}");
                    var resp = ConnectToSQL(GetConnectionString(), $@"IF EXISTS (select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa='{group}' and [CDN].[TwrGrupaPelnaNazwa] (TwG_GIDNumer) like '%Kwiaty Cięte/%') 
                    begin  
                    select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Nazwa='{group}' and [CDN].[TwrGrupaPelnaNazwa] (TwG_GIDNumer) like '%Kwiaty Cięte/%'
                    end 
                    else 
                    begin 
                    select top 1 CONCAT(TwG_GIDNumer,';', TwG_GIDTyp,';', TwG_GIDLp,';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Kod='{group}' and [CDN].[TwrGrupaPelnaNazwa] (TwG_GIDNumer) like '%Kwiaty Cięte/%' 
                    end");
                    LogAppendValue(ColorTranslator.FromHtml("#03A062"), Color.Black, $"Odnalezione GID to {resp}");

                    string[] completeGrupa = resp.Split(';');
                    if (completeGrupa.Length > 0 && group != "")
                    {
                        xLTowarInfo.TwGNumer = Int32.Parse(completeGrupa[0]);
                        xLTowarInfo.TwGTyp = Int32.Parse(completeGrupa[1]);
                        xLTowarInfo.TwGLp = Int32.Parse(completeGrupa[2]);
                        xLTowarInfo.TwGFirma = Int32.Parse(completeGrupa[3]);
                    }
                    else
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Błąd Nie znaleziono czegos dla towaru {xLTowarInfo.Nazwa}, odpowiedz to {resp}");
                        xLTowarInfo.TwGNumer = 26;
                        xLTowarInfo.TwGTyp = -16;
                        xLTowarInfo.TwGLp = 0;
                        xLTowarInfo.TwGFirma = 1403393;
                    }

                }
                catch (Exception e)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} dodawania grupy towarowej {e} " + "\n");
                }
            }
            else xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE";
            try
            {
                return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
                


            }
            catch (Exception e)
            {
                throw new Exception("Wystapil wyjatek przy dodawaniu towaru. Kod API: " + return_value + ", Wyjątek: " + e);

            }
            if (return_value != 0)
            {
                if (return_value != 1) LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} dodania towaru {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");
                LogAppendValue(Color.Black, Color.White, $"Poprwka wprowadzania grupy dla towaru {nazwa}");
                xLTowarInfo.TwrGrupa = "KWIATY DEDYKOWANE";
                return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
                string group = ConnectToSQL(connection_string, "select top 1 CONCAT(TwG_GIDNumer, ';', TwG_GIDTyp, ';', TwG_GIDLp, ';', TwG_GIDFirma) from cdn.TwrGrupy where TwG_Kod = 'KWIATY DEDYKOWANE'");
                if (return_value != 0)
                {
                    LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} poprawki x2 towaru <Z nazwa> {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                    try
                    {
                        int[] spltGrp = group.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                           .Select(x => int.Parse(x))
                           .ToArray();
                        if (spltGrp.Length >= 3)
                        {
                            xLTowarInfo.TwGNumer = spltGrp[0];
                            xLTowarInfo.TwGTyp = spltGrp[1];
                            xLTowarInfo.TwGLp = spltGrp[2];
                            xLTowarInfo.TwGFirma = spltGrp[3];
                        }
                        return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
                        if (return_value != 0) LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} poprawki x3 towaru <z GID> {nazwa}: {xLTowarInfo.Kod} z gid {xLTowarInfo.GIDNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                    }
                    catch (Exception)
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $"Błąd {++errorCount} pobierania GID grupy KWIATY DEDYKOWANE: {xLTowarInfo.Kod} z TwGNumer {xLTowarInfo.GIDNumer} TwGNumer {xLTowarInfo.TwGNumer} {'\n'} Przy grupie {xLTowarInfo.TwrGrupa} Kod API: {return_value}");
                    }
                }


            }

            XLZamkniecieTowaruInfo_20231 xLZamkniecieTowaruInfo_20231 = new XLZamkniecieTowaruInfo_20231
            {
                Wersja = XLWersjaAPI,
                Tryb = return_value != 0 ? -1 : 0,
            };
            return_value = cdn_api.cdn_api.XLZamknijTowar(XLtowar_id, xLZamkniecieTowaruInfo_20231);
            
            Update("UPDATE  c SET TwC_TcnId = 2, [TwC_DataOd]=986562792 FROM [CDN].[TwrCeny] as c join cdn.TwrCenyNag on TwC_TcnId=TCN_Id WHERE TwC_TwrTyp = 16 AND TwC_TwrFirma = " + xLTowarInfo.GIDFirma + " AND TwC_TwrNumer = " + xLTowarInfo.GIDNumer + " and TCN_RodzajCeny=2");
            Update("UPDATE  c SET TwC_TcnId = 1, [TwC_DataOd]=986562792 FROM [CDN].[TwrCeny] as c join cdn.TwrCenyNag on TwC_TcnId=TCN_Id WHERE TwC_TwrTyp = 16 AND TwC_TwrFirma = " + xLTowarInfo.GIDFirma + " AND TwC_TwrNumer = " + xLTowarInfo.GIDNumer + " and TCN_RodzajCeny=1");

            if (return_value != 0)
            {
                LogAppendValue(Color.White, Color.DarkRed, $"Błąd przy zamknięciu towaru towar : {xLTowarInfo.Kod} dodany do magazynu o gid {xLTowarInfo.GIDNumer} Kod API: {return_value}");

            }
            else
            {
                LogAppendValue(Color.Black, Color.White, $@"Utworzono towar z nazwą {xLTowarInfo.Nazwa} o kodzie {xLTowarInfo.Kod}");
                if (is_tf == true)
                {
                    string resp = ConnectToSQL(connection_string, $@"update t 
                    set  Twr_ESklep = 1 
                    from cdn.TwrKarty as t 
                    where t.Twr_GIDNumer = {xLTowarInfo.GIDNumer} and t.Twr_GIDTyp={xLTowarInfo.GIDTyp} and t.Twr_GIDFirma={xLTowarInfo.GIDFirma}");
                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do e-Sklepu {"\n"} {resp}");
                    }

                    resp = ConnectToSQL(connection_string, $@"update cdn.TwrAplikacje set TAP_Status=1, TAP_Dostepnosc=1, [TAP_Flagi] = 66 ,
                    TAP_GrupaeSklep=(select top 1 TwG_GrONumer from cdn.TwrGrupy where TwG_GIDNumer={xLTowarInfo.GIDNumer} and TwG_GIDTyp ={xLTowarInfo.GIDTyp} order by TwG_SyncId desc)  where TAP_ObiNumer = {xLTowarInfo.GIDNumer} and TAP_ObiTyp={xLTowarInfo.GIDTyp}");
                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do TwrAplikacje {"\n"} {resp}");
                    }

                    

                }
               /* else
                {
                    string resp = ConnectToSQL(connection_string, $@"update t 
                    set Twr_MobSpr = 1  
                    from cdn.TwrKarty as t 
                    where t.Twr_GIDNumer = {xLTowarInfo.GIDNumer} and t.Twr_GIDTyp={xLTowarInfo.GIDTyp} and t.Twr_GIDFirma={xLTowarInfo.GIDFirma}");

                    if (resp != "")
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} przy nadawaniu atrybutów do e-Sklepu {"\n"} {resp}");
                    }
                }*/

                string msg = ConnectToSQL(connection_string, "update cdn.TwrCeny set TwC_Wartosc=Cast('" + Regex.Replace(order.Price, "[^0-9.,]+", "", RegexOptions.Compiled).Replace(",", ".") + "' as decimal(18,4)), TwC_Marza='0.00' where TwC_TwrNumer = (select Twr_GIDNumer from cdn.TwrKarty where Twr_Kod = '" + xLTowarInfo.Kod + "') and TwC_TwrLp=1");
                if (msg != "")
                {
                    LogAppendValue(Color.White, Color.DarkRed, $"błąd przy aktualizacji ceny {msg}!");
                }
                msg = ConnectToSQL(connection_string, $@"update cdn.TwrKarty set TwR_EAN='{order.Barcode}' , Twr_MobSpr = 1 where Twr_GIDNumer ={xLTowarInfo.GIDNumer} and Twr_GIDtyp={xLTowarInfo.GIDTyp}");
                if (msg != "")
                {
                    LogAppendValue(Color.White, Color.DarkRed, $"Błąd przy aktualizacji kodu EAN (wiodące zera z przodu) {msg}!");
                }

                if (return_value == 0)
                {
                    try
                    {
                        //{ "VBN", "Klient", "Hodowca2", "S1", "S2", "S3", "IMG_URL" }
                        //Dodawanie Atrybutów
                        string flowerType = ConnectToSQL(connection_string, $"select top 1 Typ from FlowerType where KodCN={order.Code.PadLeft(8, '0')}");
                        LogAppendValue(Color.Black, Color.White, $"Rodzaj towaru dla Kodu CN {order.Code.PadLeft(8, '0')} to  {flowerType}");
                        var atr_resp = AddAtribute(new string[8] { order.Articleno, order.Grower, order.Clientnumber, order.Mark1, order.Mark2, "", order.Mark3, order.Picture }, new int[4] { xLTowarInfo.GIDFirma, xLTowarInfo.GIDLp, xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp }, flowerType);
                        if (atr_resp.Length > 0)
                        {
                            LogAppendValue(Color.White, Color.DarkBlue, $"Bład SQL przy dodawaniu atrybutów {atr_resp}");
                        }
                        else
                        {
                            LogAppendValue(Color.Black, Color.White, $"Pomyślnie dodano atrybuty");
                        }

                    }
                    catch (Exception e)
                    {
                        LogAppendValue(Color.Black, Color.White, "Nie udało się załadować obrazu " + e);
                    }
                    try
                    {
                        if (order.Picture.Length > 0)
                        {
                            // LogAppendValue(Color.DeepPink, Color.White, "Dodawanie Picture");
                            var img = DownloadImage(order.Picture);
                            if (img != null)
                            {
                                UploadImage(xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp, ((img.Size.Width > 512 && img.Size.Height > 512) ? ImageToByte2(ResizeImage(img, 512, 512)) : ImageToByte2(img)), order.Article, order.Articleno);
                            }
                            //Zmienianie Atrybutów
                            string atr_Resp = ConnectToSQL(connection_string,
                                $@"update cdn.DaneObiekty set DAO_MobSpr = 1 
                                    where DAO_ObiNumer = (select Twr_GIDNumer from cdn.TwrKarty where Twr_Kod = '{xLTowarInfo.Kod}')");
                            if (atr_Resp != "")
                            {
                                LogAppendValue(Color.White, Color.DarkBlue, "Nie udało się dodać atrybutu DAO_mobspr, odp. SQL to " + atr_Resp);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogAppendValue(Color.White, Color.DarkRed, $"Błąd nr {++errorCount}Nie udało się dodać obrazu " + ex);
                    }
                }

                if (ZZ_knt[0] > 0)
                {
                    string a = ConnectToSQL(connection_string, $@"declare @nip varchar(10)
                    set @nip = (select top 1 Knt_NipPrefiks from cdn.KntKarty where Knt_GIDNumer={ZZ_knt[0]} order by Knt_GIDNumer desc)
                    if @nip is not null 
                    begin 
                    update cdn.TwrKarty set Twr_KrajPoch=@nip where Twr_GIDNumer={xLTowarInfo.GIDNumer} and Twr_GIDLp={xLTowarInfo.GIDLp} and Twr_GIDTyp={xLTowarInfo.GIDTyp} and Twr_GIDFirma={xLTowarInfo.GIDFirma} end");
                    if (a != "") LogAppendValue(Color.White, Color.DarkRed, $@"Wystąpił błąd nr {++errorCount} dodawania kraju {a}");
                }
            }
          

            return new int[10] { xLTowarInfo.GIDNumer, xLTowarInfo.GIDTyp, xLTowarInfo.GIDLp, xLTowarInfo.GIDFirma , 0, 0, 0, 0, 0, 0};
        }
        string FindGroup(string FullAtricle)
        {
            string query = "select top 1 Grupa from dbo.FlowerGroup where NazwaKwiatu=";
            string word = "";
            string groupKey = "";
            try
            {
                string[] ArticleArr = FullAtricle.Split(' ');
               
                groupKey = "";
                        
                query += "'";
                if(ArticleArr[0]=="R")
                    query += ArticleArr[0] + " "+ ArticleArr[1];

                else
                    query += ArticleArr[0];
                           
                word = ConnectToSQL(connection_string, query+"' order by len(Grupa) desc");
                LogAppendValue(Color.DeepPink, Color.White, query+"\nWynik: "+word);
            }
            catch (Exception e)
            {
                LogAppendValue(Color.White, Color.DarkRed, $@"Błąd tworzenia zapytania o grupę towarową {e}");
                return word;
            }
            return word;
        }

        /*string FindGroup(string FullAtricle)
        {
            string query = "select top 1 Grupa from dbo.FlowerGroup where NazwaKwiatu=";
            string word = "";
            string groupKey = "";
            try
            {
                string[] ArticleArr = FullAtricle.Split(' ');
                for (int i = 0; i < ArticleArr.Count(); i++)
                {
                    for (int j = 0; j < ArticleArr.Count() - i; j++)
                    {
                        groupKey = "";
                        for (int k = 0; k <= i; k++)
                        {

                            if (k == 0)
                            {
                                query += "'";
                                groupKey += ArticleArr[k + j];
                            }
                            else groupKey += " " + ArticleArr[k + j];
                             if (k == i)
                            {
                              if (groupKey.Length > 1)
                             {
                                 query += groupKey + "'";
                            }
                             else query += " '";

                             }
                             }
                             if (i != ArticleArr.Count() - 1 && j != (ArticleArr.Count() - i)) query += " or NazwaKwiatu=";
                        }
                    }
                    word = ConnectToSQL(connection_string, query + " order by len(Grupa) desc");
                    LogAppendValue(Color.DeepPink, Color.White, query + "\nWynik: " + word);
                }
            catch (Exception e)
            {
                LogAppendValue(Color.White, Color.DarkRed, $@"Błąd tworzenia zapytania o grupę towarową {e}");
                return word;
            }
            return word;
        }*/
        public byte[] ImageToByte2(Image img)
        {
            //LogAppendValue(Color.DeepPink, Color.White, "Wchodzę do Image To Byte 2");
            try
            {
                using (var stream = new MemoryStream())
                {
                    img.Save(stream, ImageFormat.Jpeg);
                    //LogAppendValue(Color.DeepPink, Color.White, "Wychodzę z Image To Byte 2");
                    return stream.ToArray();
                }
            }
            catch(Exception e)
            {
                //LogAppendValue(Color.DeepPink , Color.White, $@"Błąd ImgageToByte2 {e}");
                return null;
            }
        }
        Image DownloadImage(string url)
        {
            //LogAppendValue(Color.DeepPink, Color.White, "Wchodzę do Download Image");
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                using (WebClient web = new WebClient())
                {
                    using(System.IO.Stream stream = web.OpenRead(url))
                    {
                        //LogAppendValue(Color.DeepPink, Color.White, "Wychodzę z Download Image ");
                        return Image.FromStream(stream);
                    }
                   /* WebRequest requestPic = WebRequest.Create(url);

                    WebResponse responsePic = requestPic.GetResponse();

                    LogAppendValue(Color.DeepPink, Color.White, "Wychodzę z Download Image ");
                    return Image.FromStream(responsePic.GetResponseStream());*/

                }
            }
            catch(Exception e)
            {
                LogAppendValue(Color.DeepPink, Color.White, $@"Błąd Download Image {e}");
                return null;
            }

        }
        private Bitmap LoadImage(string address)
        {
           // LogAppendValue(Color.DeepPink, Color.White, "Wchodzę do Load Image");
            AttachThreadToClarion(true);
            try
            {
                var request = WebRequest.Create(address);
                Image image;
                using (var response = request.GetResponse())
                using (var stream = response.GetResponseStream())
                {
                    image = Image.FromStream(stream);
                }
                //LogAppendValue(Color.DeepPink, Color.White, "Wychodzę z Load Image");
                return (Bitmap)image;
            }
            catch(Exception e)
            {
                LogAppendValue(Color.DeepPink, Color.White, $@"Błąd Load Image {e}");
                return null;
            }
        }
        public Bitmap ResizeImage(Image image, int width, int height)
        {
           // LogAppendValue(Color.DeepPink, Color.White, "Wchodzę do Resize Image");
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            try
            {
                destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

                using (var graphics = Graphics.FromImage(destImage))
                {
                    graphics.CompositingMode = CompositingMode.SourceCopy;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                    using (var wrapMode = new ImageAttributes())
                    {
                        wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                        graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                    }
                }
            }
            catch (Exception e)
            {
                LogAppendValue(Color.DeepPink, Color.White, $@"Błąd Resize Image {e}");
               
            }

            //LogAppendValue(Color.DeepPink, Color.White, "Wychodzę z Resize Image ");
            return destImage;
        }
        private void UploadImage(int GIDnumer, int GIDtyp, byte[] image, string nazwa, string kod)
        {
            //LogAppendValue(Color.DeepPink, Color.White, "Wchodzę do Upload Image");
            try
            {
                string query_part_1 = $@"SET IDENTITY_INSERT CDN.DaneBinarne ON 
            declare @file_stream VARCHAR(MAX) 
            set @file_stream='' 
            INSERT INTO CDN.DaneBinarne 
	        ( 
	            DAB_Nazwa
	            , DAB_TypID
	            , DAB_Rozszerzenie
	            , DAB_Rozmiar
	            , DAB_Dane
	            , DAB_Kod
	            , DAB_Czasmodyfikacji
	            , DAB_OpeNumer
	            , DAB_Usuwac
	            , DAB_Archiwalny
	            , DAB_Tlumaczenie
	            , DAB_Jezyk
	            , DAB_Aktywny
	            , DAB_CzasArchiwizacji
	            , DAB_OpisArchiwizacji 
	            , DAB_PPPrawa
	            , DAB_PKPrawa
	            , DAB_eSklep
	            , DAB_ID
	        ) 
	        Values
                ( 
		        '{kod}'
		        , 326
		        , 'JPG'
		        , 0
		        , {"0x" + NewConverter(image)}
		        , '{kod}'
		        , 681390307
		        , 1
		        , 0
		        , 0
		        , (SELECT MAX(DAB_Tlumaczenie)+1 FROM CDN.DaneBinarne)
		        , 0
		        , 0
		        , 2000000000
		        , ''
		        , 0
		        , 0
		        , 1
		        , (SELECT ISNULL(MAX(DAB_ID)+1,1) FROM CDN.DaneBinarne)
		        )
            SET IDENTITY_INSERT CDN.DaneBinarne OFF";
               // LogAppendValue(Color.DeepPink, Color.White, query_part_1);
                

                // richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,query_part_1);
                string msg = ConnectToSQL(connection_string, query_part_1);
                //richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,"upload obrazeczka: " + msg);
                //LogAppendValue(Color.DeepPink, Color.White, "upload obrazeczka: " + msg);

                //Powiązanie obrazu z towarem
                if (msg.Length <= 1)
                {
                    query_part_1 = $@"declare @dabid as int
                set @dabid = (select MAX(DAB_ID) from cdn.DaneBinarne where DAB_Kod='{kod}')
                INSERT INTO CDN.DaneObiekty
	                (
	                DAO_DABId,
	                DAO_ObiTyp,
	                DAO_ObiNumer,
	                DAO_ObiLp,
	                DAO_ObiSubLp,
	                DAO_Domyslna,
	                DAO_Blokada,
	                DAO_PPPrawa,
	                DAO_PKPrawa,
	                DAO_eSklep,
	                DAO_iMall,
	                DAO_MobSpr,
	                DAO_BI,
	                DAO_Systemowa,
	                DAO_Pozycja,
	                DAO_Retail,
	                DAO_WMSZarzadzanie,
	                DAO_WMSMagazynier,
	                DAO_CzasModyfikacji)
	                Values
	                (@dabid,{GIDtyp},{GIDnumer},0,0,0,0,0,0,1,1,0,1,0,1,0,0,0,(select cdn.DateToTS(GETDATE())))
                ";
                    //LogAppendValue(Color.DeepPink, Color.White, query_part_1);
                    //richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,query_part_1);
                    msg = ConnectToSQL(connection_string, query_part_1);
                    //richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,"upload obrazeczka: " + msg);
                    
                    //LogAppendValue(Color.DeepPink, Color.White, "upload obrazeczka: " + msg);

                }
            }
            catch (Exception e)
            {
                LogAppendValue(Color.DeepPink, Color.White, $@"Błąd Upload Image {e}");

            }


        }
        public string NewConverter(byte[] data)
        {
            if (data == null)
            {
                throw new ArgumentNullException("data");
            }

            int length = data.Length;
            char[] hex = new char[length * 2];
            int num1 = 0;
            for (int index = 0; index < length * 2; index += 2)
            {
                byte num2 = data[num1++];
                hex[index] = GetHexValue(num2 / 0x10);
                hex[index + 1] = GetHexValue(num2 % 0x10);
            }
            return new string(hex);
        }
        private char GetHexValue(int i)
        {
            if (i < 10)
            {
                return (char)(i + 0x30);
            }
            return (char)(i - 10 + 0x41);
        }

        public void LogAppendValue( Color text_color, Color background_color, string text)
        {
            try
            {
                infoLog.Add(Tuple.Create(text_color, background_color, text));
                tw2.WriteLine(text);
            }
            catch (Exception)
            {
                
            }
        }

        public string AddAtribute(string[] WartoscAtrybutu, int[] dokInfo, string flowerType)
        {
            string ogrodnik = $@"if not exists (select 1 from cdn.Atrybuty left join cdn.AtrybutyKlasy on AtK_ID=Atr_AtkId where Atr_ObiTyp = {dokInfo[3]} AND Atr_ObiNumer = {dokInfo[2]} AND Atr_ObiLp = 0 AND Atr_ObiSubLp = 0 AND AtK_Nazwa = 'VBN') begin 
            insert into cdn.Atrybuty(Atr_ObiTyp, Atr_ObiFirma, Atr_ObiNumer, Atr_ObiLp, Atr_ObiSubLp, Atr_AtkId, Atr_Wartosc, Atr_AtrTyp, Atr_AtrFirma, Atr_AtrNumer, Atr_AtrLp, Atr_AtrSubLp, Atr_OptimaId, Atr_Grupujacy, Atr_Pozycja) 
                                VALUES({dokInfo[3]}, 
                                {dokInfo[0]}, 
                                {dokInfo[2]}, 
                                0, 
                                0, (Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa = 'VBN'), '{WartoscAtrybutu[0]}', 0, 0, 0, 0, 0, 0, 0, 1); end;";
            string resp = ConnectToSQL(connection_string, ogrodnik);
            //if(resp!="") richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,$@"Dla VBN {WartoscAtrybutu[0]} wystąpił błąd: " + resp);
            ogrodnik = $@"declare @resp as VARCHAR(MAX)
            declare @grower as VARCHAR(MAX)
            set @grower='{WartoscAtrybutu[1]}'
            if not exists (select 1 from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower) 
                        begin 
			            insert into cdn.AtrybutyWartosci(AtW_AtKId, AtW_Wartosc, AtW_OddZrdId, AtW_Lp) 
                        Values( (select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2'), @grower, 0, 0); 
			            set @resp = Right('00000' + Convert(varchar(5),(select top 1 AtW_ID from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower)),5) 
			            end
			            else
			            begin
                        set @resp = Right('00000' + Convert(varchar(5),(select top 1 AtW_ID from cdn.AtrybutyWartosci where AtW_AtKId=(select top 1 AtK_ID from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atr_AtkId = AtK_ID where AtK_Nazwa='Hodowca2') and AtW_Wartosc=@grower)),5)
			            end
			            if not exists (select 1 from cdn.Atrybuty right join cdn.AtrybutyKlasy on AtK_ID=Atr_AtkId where Atr_ObiTyp = {dokInfo[3]} AND Atr_ObiNumer = {dokInfo[2]} AND Atr_ObiLp = 0 AND Atr_ObiSubLp = 0 AND AtK_Nazwa = 'Hodowca2') begin 
                        insert into cdn.Atrybuty(Atr_ObiTyp, Atr_ObiFirma, Atr_ObiNumer, Atr_ObiLp, Atr_ObiSubLp, Atr_AtkId, Atr_Wartosc, Atr_AtrTyp, Atr_AtrFirma, Atr_AtrNumer, Atr_AtrLp, Atr_AtrSubLp, Atr_OptimaId, Atr_Grupujacy, Atr_Pozycja) 
                                VALUES({dokInfo[3]}, 
                                {dokInfo[0]}, 
                                {dokInfo[2]}, 
                                0, 
                                0, (Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa = 'Hodowca2'), @grower, 0, 0, 0, 0, 0, 0, 0, 1); end";
            resp = ConnectToSQL(connection_string, ogrodnik);
            //if (resp != "") richTextBox_Log_write(newLog.richTextBox1, Color.Black, Color.White,$@"Dla hodowcy {WartoscAtrybutu[1]} wystąpił błąd: " + ogrodnik);

            string ifnotexists = "if not exists (select 1 from cdn.Atrybuty left join cdn.AtrybutyKlasy on AtK_ID=Atr_AtkId where Atr_ObiTyp = 960 AND Atr_ObiNumer = " + dokInfo[2].ToString() + " AND Atr_ObiLp = 0 AND Atr_ObiSubLp = 0 AND (AtK_Nazwa = '" + KlasyAtrybutuNazwa.LastOrDefault() + "'";
            ifnotexists += ")) begin ";
            string query = ifnotexists + "insert into cdn.Atrybuty(Atr_ObiTyp, Atr_ObiFirma, Atr_ObiNumer, Atr_ObiLp, Atr_ObiSubLp, Atr_AtkId, Atr_Wartosc, Atr_AtrTyp, Atr_AtrFirma, Atr_AtrNumer, Atr_AtrLp, Atr_AtrSubLp, Atr_OptimaId, Atr_Grupujacy, Atr_Pozycja) " +
                                "VALUES(" + dokInfo[3].ToString() + ", " +
                                dokInfo[0].ToString() + ", " +
                                dokInfo[2].ToString() + ", " +
                                "0, " +
                                "0," + "(Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='" + KlasyAtrybutuNazwa.LastOrDefault() + "')" + ",'" + WartoscAtrybutu.LastOrDefault() + "', 0, 0, 0, 0, 0, 0, 0,1)"; 
            query += "; end";
            resp = ConnectToSQL(connection_string, query); 
            //S1 S2 S3 
            for (int i = 3; i < WartoscAtrybutu.Length-1; i++)
            {
                resp += ConnectToSQL(connection_string, $@"update cdn.Atrybuty set Atr_Wartosc = '{WartoscAtrybutu[i]}' where Atr_ObiNumer = {dokInfo[2]} and Atr_ObiTyp = {dokInfo[3]} and Atr_ObiFirma = {dokInfo[0]} and Atr_AtkId=(Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='{KlasyAtrybutuNazwa[i]}')");
            } 

            //02.02.2022 Atrybuty na esklep
            string[] eSklepAtrybuty = new string[4] { WartoscAtrybutu[3], WartoscAtrybutu[4], WartoscAtrybutu[5], WartoscAtrybutu[6]}; 
            switch (flowerType)
            {
                case "Doniczkowe":
                    for (int i = 0; i < KlasyAtrybutuEsklepDoniczkowe.Length; i++)
                    {
                        resp += ConnectToSQL(connection_string, $@"update cdn.AtrybutyHist set AtH_Wartosc = '{eSklepAtrybuty[i]}' where AtH_ObiNumer = {dokInfo[2]} and AtH_ObiTyp = {dokInfo[3]} and AtH_ObiFirma = {dokInfo[0]} and AtH_AtkId=(Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='{KlasyAtrybutuEsklepDoniczkowe[i]}')");
                    }
                    break;
                default:
                    for (int i = 0; i < KlasyAtrybutuEsklepCiete.Length; i++)
                    {
                        resp += ConnectToSQL(connection_string, $@"update cdn.AtrybutyHist set AtH_Wartosc = '{eSklepAtrybuty[i]}' where AtH_ObiNumer = {dokInfo[2]} and AtH_ObiTyp = {dokInfo[3]} and AtH_ObiFirma = {dokInfo[0]} and AtH_AtkId=(Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='{KlasyAtrybutuEsklepCiete[i]}')");
                    }
                    break;
            }

            return resp;
        }
        public string ChangeDate(int clarionDate, DateTime dateTime, int gID)
        {
            var query = $@"declare @zamNumer int 
                declare @pomNumer int 
                set @zamNumer = (select MAX(ZaN_ZamNumer)+1 from cdn.ZamNag where ZaN_ZamMiesiac={dateTime.Month}) 
                set @pomNumer = (select MAX(ZaN_PomNumer)+1 from cdn.ZamNag where ZaN_ZamMiesiac={dateTime.Month}) 
                update cdn.ZamNag set ZaN_ZamMiesiac={dateTime.Month}, 
                ZaN_ZamNumer=(select ISNULL(@zamNumer,1)), 
                ZaN_PomNumer=(select ISNULL(@zamNumer,1)), 
                ZaN_PomMiesiac=3, ZaN_DataWystawienia={clarionDate}, ZaN_DataRealizacji={clarionDate}, 
                Zan_RokMiesiac={dateTime.ToString("yyyyMM")}  
                where ZaN_GIDNumer={gID}";
            return ConnectToSQL(connection_string,query);
        }

        public string ChangeDateZS(int clarionDate, DateTime dateTime, int gID)
        {
            var query = $@"declare @zamNumer int 
                declare @pomNumer int 
                set @zamNumer = (select MAX(ZaN_ZamNumer)+1 from cdn.ZamNag where ZaN_ZamMiesiac={dateTime.Month}) 
                set @pomNumer = (select MAX(ZaN_PomNumer)+1 from cdn.ZamNag where ZaN_ZamMiesiac={dateTime.Month}) 
                update cdn.ZamNag set ZaN_ZamMiesiac={dateTime.Month}, 
                ZaN_ZamNumer=(select ISNULL(@zamNumer,1)), 
                ZaN_PomNumer=(select ISNULL(@zamNumer,1)), 
                ZaN_PomMiesiac=3, ZaN_DataWystawienia={GetClarionDate(dateTime)}, ZaN_DataRealizacji={GetClarionDate(dateTime)}, ZaN_DataWaznosci={GetClarionDate(dateTime)}, 
                Zan_RokMiesiac={dateTime.ToString("yyyyMM")}  
                where ZaN_GIDNumer={gID}";
            ConnectToSQL(connection_string, query);

            var query2 = $@"UPDATE CDN.ZamElem set ZaE_DataWaznosciRez = {GetClarionDate(dateTime)} where ZaE_GIDNumer={gID}";
            
            return ConnectToSQL(connection_string, query2);
        }
        public int XLLogin()
        {
            AttachThreadToClarion(true);
            return cdn_api.cdn_api.XLLogin(new XLLoginInfo_20231()
            {
                Wersja = XLWersjaAPI,
                ProgramID = "ONE_S_CSV_To_ZZ",
                //PlikLog = (Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\api.log"),
                //PlikLog = (@"C:\OneSupport\XML\api.log"),
                Baza = XLDatabase,
                OpeIdent = XLUser,
                OpeHaslo = XLPass,
                TrybWsadowy = XLTrybWsadowy,
                // TrybNaprawy = 0,
                //SerwerKlucza=XLKey,
                //PortHTTPSerweraKlucza=1,
                UtworzWlasnaSesje = XLUtworzWlasnaSesje,
                Winieta = XLWinieta,
            }, ref sesja);
        }
        public int XLLogout()
        {
            AttachThreadToClarion(true);
            int str = -1;
            XLLogoutInfo_20231 xLLogoutInfo_20231 = new XLLogoutInfo_20231()
            {
                Wersja = 20202,
                //TrybUsuwaniaDok = 1,
            };
            str = cdn_api.cdn_api.XLLogoutEx(xLLogoutInfo_20231, sesja);
            return str;

        }

        public string GetConnectionString()
        {
            XLPolaczenieInfo_20231 polaczenieInfo = new XLPolaczenieInfo_20231 { Wersja = 20231 };
            cdn_api.cdn_api.XLPolaczenie(polaczenieInfo);
            //connection_string = "Server=" + polaczenieInfo.Serwer + ";Database=" + polaczenieInfo.Baza + ";User Id=" + XLUser + ";Password=" + XLPass;
            connection_string = "Server=" + polaczenieInfo.Serwer + ";Database=" + polaczenieInfo.Baza + XLConString; //";User ID=SA;Password=DoMarSQL17";
            return connection_string;
        }
        public string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "błąd SQL:" + ex.ToString() + "\n";
            }
            return str;
        }
        public List<int> ConnectListToSQL(string connection_string, string query)
        {
            List<int> columnData = new List<int>();

            using (SqlConnection connection = new SqlConnection(connection_string))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            columnData.Add(int.Parse(reader.GetString(0)));
                            columnData.Add(int.Parse(reader.GetString(1)));
                            columnData.Add(int.Parse(reader.GetString(2)));
                            columnData.Add(int.Parse(reader.GetString(3)));
                        }
                    }
                }
            }
            return columnData;
        }

        public List<int> ConnectListToSQLZZ(string connection_string, string query)
        {
            List<int> columnData = new List<int>();

            using (SqlConnection connection = new SqlConnection(connection_string))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            columnData.Add(int.Parse(reader.GetString(0)));
                            columnData.Add(int.Parse(reader.GetString(1)));
                            columnData.Add(int.Parse(reader.GetString(2)));
                            columnData.Add(int.Parse(reader.GetString(3)));
                            columnData.Add(int.Parse(reader.GetString(4)));
                            columnData.Add(int.Parse(reader.GetString(5)));
                            columnData.Add(int.Parse(reader.GetString(6)));
                            columnData.Add(int.Parse(reader.GetString(7)));
                            columnData.Add(int.Parse(reader.GetString(8)));
                            columnData.Add(int.Parse(reader.GetString(9)));
                            
                        }
                    }
                }
            }
            return columnData;
        }
        public List<string> ConnectListStrToSQL(string connection_string, string query)
        {
            List<string> columnData = new List<string>();

            using (SqlConnection connection = new SqlConnection(connection_string))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            columnData.Add(reader.GetString(0));
                            columnData.Add(reader.GetString(1));
                            columnData.Add(reader.GetString(2)); 
                        }
                    }
                }
            }
            return columnData;
        }

        public string Update(string sql)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.UpdateCommand = new SqlCommand(sql, connection);
                    adapter.UpdateCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row updated !! ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sql;
        }

        public int GetClarionDate(DateTime dotNetDate)
        {
            CultureInfo ci = new CultureInfo("pl-PL");
            DateTime baseDate = DateTime.Parse("01.01.1801", ci);
            TimeSpan ts = dotNetDate - baseDate;
            int l = ts.Days + 4;
            return l;
        }
        public int GetClarionTime(DateTime time)
        {
            int l = time.Hour * 360000;
            l += time.Minute * 6000;
            l += time.Second * 100;
            l += time.Millisecond / 10;
            return l + 1;
        }
        public string Truncate(string value, int maxChars)
        {
            return value.Length <= maxChars ? value : value.Substring(0, maxChars);
        }
        public int zwrocTryb()
        {
            return RodzajDzialania;
        }
    public List<Tuple<Color, Color, string>> zwrocLog()
        {
            return infoLog;
        }
}
}

