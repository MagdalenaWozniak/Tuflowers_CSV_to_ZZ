
namespace Tufl_CSV_To_ZZ
{
    partial class Decyzja
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Decyzja));
            this.label1 = new System.Windows.Forms.Label();
            this.buttonTak = new System.Windows.Forms.Button();
            this.buttonNie = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label1.Location = new System.Drawing.Point(12, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(354, 23);
            this.label1.TabIndex = 2;
            this.label1.Text = "Taki numer dostawy już istnieje. Czy chcesz kontynuować?";
            // 
            // buttonTak
            // 
            this.buttonTak.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.buttonTak.Location = new System.Drawing.Point(204, 67);
            this.buttonTak.Name = "buttonTak";
            this.buttonTak.Size = new System.Drawing.Size(79, 31);
            this.buttonTak.TabIndex = 0;
            this.buttonTak.Text = "TAK";
            this.buttonTak.UseVisualStyleBackColor = false;
            this.buttonTak.Click += new System.EventHandler(this.buttonTak_Click);
            // 
            // buttonNie
            // 
            this.buttonNie.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.buttonNie.Location = new System.Drawing.Point(56, 67);
            this.buttonNie.Name = "buttonNie";
            this.buttonNie.Size = new System.Drawing.Size(82, 31);
            this.buttonNie.TabIndex = 1;
            this.buttonNie.Text = "NIE";
            this.buttonNie.UseVisualStyleBackColor = false;
            this.buttonNie.Click += new System.EventHandler(this.buttonNie_Click);
            // 
            // Decyzja
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(368, 130);
            this.Controls.Add(this.buttonTak);
            this.Controls.Add(this.buttonNie);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Decyzja";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ostrzeżenie";
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonTak;
        private System.Windows.Forms.Button buttonNie;
        #endregion
    }
}