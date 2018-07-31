namespace ExcelMate
{
    partial class RoundSelector
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RoundSelector));
            this.bnCancel = new System.Windows.Forms.Button();
            this.bnOverWrite = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.cbOverwrite = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // bnCancel
            // 
            this.bnCancel.Location = new System.Drawing.Point(8, 41);
            this.bnCancel.Name = "bnCancel";
            this.bnCancel.Size = new System.Drawing.Size(75, 23);
            this.bnCancel.TabIndex = 0;
            this.bnCancel.Text = "Cancel";
            this.bnCancel.UseVisualStyleBackColor = true;
            this.bnCancel.Click += new System.EventHandler(this.bnCancel_Click);
            // 
            // bnOverWrite
            // 
            this.bnOverWrite.Location = new System.Drawing.Point(220, 41);
            this.bnOverWrite.Name = "bnOverWrite";
            this.bnOverWrite.Size = new System.Drawing.Size(74, 23);
            this.bnOverWrite.TabIndex = 1;
            this.bnOverWrite.Text = "Overwrite";
            this.bnOverWrite.UseVisualStyleBackColor = true;
            this.bnOverWrite.Click += new System.EventHandler(this.bnOverWrite_Click);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Control;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(8, 6);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(291, 33);
            this.textBox1.TabIndex = 4;
            this.textBox1.TabStop = false;
            this.textBox1.Text = "This rider already have times in all runs. \r\nIn order to save this time we need t" +
                "o overwrite one. Which?";
            // 
            // cbOverwrite
            // 
            this.cbOverwrite.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOverwrite.FormattingEnabled = true;
            this.cbOverwrite.Items.AddRange(new object[] {
            "Select run",
            "First run",
            "Second run"});
            this.cbOverwrite.Location = new System.Drawing.Point(92, 42);
            this.cbOverwrite.Name = "cbOverwrite";
            this.cbOverwrite.Size = new System.Drawing.Size(121, 21);
            this.cbOverwrite.TabIndex = 5;
            this.cbOverwrite.SelectedIndexChanged += new System.EventHandler(this.cbOverwrite_SelectedIndexChanged);
            // 
            // RoundSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(302, 76);
            this.ControlBox = false;
            this.Controls.Add(this.cbOverwrite);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.bnOverWrite);
            this.Controls.Add(this.bnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "RoundSelector";
            this.ShowInTaskbar = false;
            this.Text = "A time already exist, cancel or overwrite?";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bnCancel;
        private System.Windows.Forms.Button bnOverWrite;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ComboBox cbOverwrite;
    }
}