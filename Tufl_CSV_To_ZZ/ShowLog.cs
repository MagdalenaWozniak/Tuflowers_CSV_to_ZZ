using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tufl_CSV_To_ZZ
{
    public partial class ShowLog : Form
    {
        Button_Handle mainClass = new Button_Handle();
        static int ErrorList = 0;

        public ShowLog()
        {
            InitializeComponent();
            //log = mainClass.zwrocLog();
            //richTextBox1.Lines = log;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label3.Text = $@"Ilość błędów: {mainClass.GetErrCount()}";
            if (mainClass.GetErrCount()>0)
            {
                label3.ForeColor = Color.DarkRed;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "txt files (*txt)|*.txt*";
            saveFileDialog.DefaultExt = "txt";
            saveFileDialog.AddExtension = true;
            saveFileDialog.RestoreDirectory = true;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                StreamWriter writer = new StreamWriter(saveFileDialog.OpenFile());
                for (int i = 0; i < richTextBox1.Lines.Length; i++)
                {
                    writer.WriteLine(richTextBox1.Lines[i]);
                }
                writer.Dispose();
                writer.Close();
            }
        }
        public void richTextBox_Log_write(List<Tuple<Color,Color,string>> logList)
        {
            try
            { 
                for (int i = 0; i < logList.Count; i++)
                {
                    int start = richTextBox1.Text.Length;
                    string rtbText = $"\n{(richTextBox1.Lines.Length + 1)}. " + logList[i].Item3;
                    richTextBox1.AppendText(rtbText);
                    richTextBox1.Select(start, rtbText.Length);
                    if (logList[i].Item1 == null)
                    {
                        richTextBox1.SelectionColor = Color.Black;
                    }
                    else richTextBox1.SelectionColor = logList[i].Item1;
                    if (logList[i].Item2 != null)
                    {
                        richTextBox1.SelectionBackColor = logList[i].Item2;
                    }
                    if (logList[i].Item2 == Color.DarkRed)
                    {
                        ErrorList++;
                    }
                }
                label2.Text = "Ilość linii tekstu" + richTextBox1.Lines.Length.ToString();
            }
            catch (Exception e)
            {
                richTextBox1.AppendText("\n" + $"Błąd zaciągania tekstu do linii {richTextBox1.Lines.Length + 1}, tekst {logList[richTextBox1.Lines.Length].Item3}, blad {e}");
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
            mainClass.LoadNewCSV();
        }

        private void ShowLog_Activated(object sender, EventArgs e)
        {
            
            try
            {
                /*
                for (int i = 0; i < log.Length; i++)
                {
                    richTextBox1.Lines[i] = log[i];
                }
                */
                switch (mainClass.zwrocTryb())
                {
                    case -1:
                        label1.Text = "Nieznany tryb";
                        break;
                    case 0:
                        label1.Text = "Domyślne ZZ->ZS";
                        break;
                    case 1:
                        label1.Text = "Tylko ZS";
                        break;
                    case 2:
                        label1.Text = "Tylko ZZ";
                        break;
                    default:
                        label1.Text = "Nieznany tryb";
                        break;
                }
                
            }
            catch (Exception er)
            {
                throw new Exception($@"Wystąpił błąd importu logu. {er}");
            }
            
            }


        private void ShowLog_FormClosed(object sender, FormClosedEventArgs e)
        {
            richTextBox1.Text = "";
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
