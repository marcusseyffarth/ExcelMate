namespace ExcelMate
{
    partial class OverWrite
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OverWrite));
            this.bnCancel = new System.Windows.Forms.Button();
            this.bnOverWrite2 = new System.Windows.Forms.Button();
            this.bnOverWrite1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // bnCancel
            // 
            this.bnCancel.Location = new System.Drawing.Point(12, 44);
            this.bnCancel.Name = "bnCancel";
            this.bnCancel.Size = new System.Drawing.Size(75, 23);
            this.bnCancel.TabIndex = 0;
            this.bnCancel.Text = "Cancel";
            this.bnCancel.UseVisualStyleBackColor = true;
            this.bnCancel.Click += new System.EventHandler(this.bnCancel_Click);
            // 
            // bnOverWrite2
            // 
            this.bnOverWrite2.Location = new System.Drawing.Point(273, 44);
            this.bnOverWrite2.Name = "bnOverWrite2";
            this.bnOverWrite2.Size = new System.Drawing.Size(90, 23);
            this.bnOverWrite2.TabIndex = 1;
            this.bnOverWrite2.Text = "Overwrite run 2";
            this.bnOverWrite2.UseVisualStyleBackColor = true;
            this.bnOverWrite2.Click += new System.EventHandler(this.bnOverWrite2_Click);
            // 
            // bnOverWrite1
            // 
            this.bnOverWrite1.Location = new System.Drawing.Point(172, 44);
            this.bnOverWrite1.Name = "bnOverWrite1";
            this.bnOverWrite1.Size = new System.Drawing.Size(95, 23);
            this.bnOverWrite1.TabIndex = 2;
            this.bnOverWrite1.Text = "Overwrite run 1";
            this.bnOverWrite1.UseVisualStyleBackColor = true;
            this.bnOverWrite1.Click += new System.EventHandler(this.bnOverWrite1_Click);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Control;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(12, 7);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(319, 31);
            this.textBox1.TabIndex = 3;
            this.textBox1.Text = "A time for the second run already exist for this rider.\r\nWhat would you like to d" +
                "o?";
            // 
            // OverWrite
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(366, 70);
            this.ControlBox = false;
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.bnOverWrite1);
            this.Controls.Add(this.bnOverWrite2);
            this.Controls.Add(this.bnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "OverWrite";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "A time for run 2 for this rider already exist in this round";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bnCancel;
        private System.Windows.Forms.Button bnOverWrite2;
        private System.Windows.Forms.Button bnOverWrite1;
        private System.Windows.Forms.TextBox textBox1;
    }
}