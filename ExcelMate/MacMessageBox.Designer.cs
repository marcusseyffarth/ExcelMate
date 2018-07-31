namespace ExcelMate
{
    partial class MacMessageBox
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MacMessageBox));
            this.label1 = new System.Windows.Forms.Label();
            this.bnOk = new System.Windows.Forms.Button();
            this.bnTwo = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.WhiteSmoke;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(12, 1);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(149, 39);
            this.label1.TabIndex = 0;
            this.label1.Text = "Connected!";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // bnOk
            // 
            this.bnOk.BackColor = System.Drawing.Color.WhiteSmoke;
            this.bnOk.ForeColor = System.Drawing.Color.Black;
            this.bnOk.Location = new System.Drawing.Point(47, 43);
            this.bnOk.Name = "bnOk";
            this.bnOk.Size = new System.Drawing.Size(75, 23);
            this.bnOk.TabIndex = 1;
            this.bnOk.Text = "Ok";
            this.bnOk.UseVisualStyleBackColor = false;
            this.bnOk.Click += new System.EventHandler(this.bnOk_Click);
            // 
            // bnTwo
            // 
            this.bnTwo.BackColor = System.Drawing.Color.WhiteSmoke;
            this.bnTwo.ForeColor = System.Drawing.Color.Black;
            this.bnTwo.Location = new System.Drawing.Point(86, 43);
            this.bnTwo.Name = "bnTwo";
            this.bnTwo.Size = new System.Drawing.Size(75, 23);
            this.bnTwo.TabIndex = 2;
            this.bnTwo.Text = "Cancel";
            this.bnTwo.UseVisualStyleBackColor = false;
            this.bnTwo.Visible = false;
            this.bnTwo.Click += new System.EventHandler(this.bnTwo_Click);
            // 
            // MacMessageBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(173, 73);
            this.Controls.Add(this.bnTwo);
            this.Controls.Add(this.bnOk);
            this.Controls.Add(this.label1);
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MacMessageBox";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button bnOk;
        private System.Windows.Forms.Button bnTwo;
    }
}