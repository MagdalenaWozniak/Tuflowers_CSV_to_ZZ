using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Tufl_CSV_To_ZZ.WyborTrybuImportu;

namespace Tufl_CSV_To_ZZ
{
    public partial class Decyzja : Form
    {
        //static Button_Handle class2 = new Button_Handle();
        //static WyborTrybuImportu import;
        DateTime data;
        string textBox;
        int index;
        string dostawa;
        Button_Handle class1;
       
        public Decyzja(DateTime data, string textBox, int index, Button_Handle class1, string dostawa)
        {
            InitializeComponent();
            this.data=data;
            this.textBox=textBox;
            this.index=index;
            this.class1 = class1;
            this.dostawa = dostawa;
            if(index==1)
            {
                label1.Text = "Taki numer dostawy już istnieje. Czy chcesz kontynuować?";
            }
            else
            {
                label1.Text = "Taki numer dostawy nie istnieje. Czy chcesz kontynuować?";
            }
            //MessageBox.Show("Index: " + this.index + "\nTryb: " + this.textBox + "\nData: " + this.data);
            
            
            
        }

        private void buttonTak_Click(object sender, EventArgs e)
        {
            
            class1.ReadVariables(data, textBox == "" ? "NULL" : textBox, index, dostawa);
            
            this.Close();
            ActiveForm.Close();
        }

        private void buttonNie_Click(object sender, EventArgs e)
        {
            
            this.Close();
        }
    }
}
