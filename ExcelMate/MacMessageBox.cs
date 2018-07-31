using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExcelMate
{
    public partial class MacMessageBox : Form
    {
        public MacMessageBox()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public MacMessageBox(String message)
        {
            InitializeComponent();
            label1.Text = message;

            Graphics g = this.CreateGraphics();
            int m_nWidth = Convert.ToInt32(g.MeasureString(message, label1.Font).Width* 1.2) + 30;

            this.Width = m_nWidth;
            label1.Width = m_nWidth;

            System.Drawing.Point l = label1.Location;
            l.X = 0;
            label1.Location = l;

            System.Drawing.Point p = bnOk.Location;
            p.X = m_nWidth / 2 - 20;
            bnOk.Location = p;
        }
       
        private void bnOk_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        /// <summary>
        /// Only 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="buttons"></param>
        public MacMessageBox(String message, MessageBoxButtons buttons)
        {
            InitializeComponent();
            label1.Text = message;

            Graphics g = this.CreateGraphics();
            int m_nWidth = Convert.ToInt32(g.MeasureString(message, label1.Font).Width* 1.2) + 30;

            this.Width = m_nWidth;
            label1.Width = m_nWidth;

            System.Drawing.Point l = label1.Location;
            l.X = 0;
            label1.Location = l;

            System.Drawing.Point p;
            switch (buttons){
                case MessageBoxButtons.OK: 
                    p = bnOk.Location;
                    p.X = m_nWidth / 2 - 20;
                    bnOk.Location = p;

                    bnTwo.Visible = false;
                    bnOk.Visible = true;
                    break;
                case MessageBoxButtons.OKCancel: 
                    p = bnOk.Location;
                    p.X = m_nWidth / 4 - 40;
                    bnOk.Location = p;

                    p = bnTwo.Location;
                    p.X = 3*(m_nWidth / 4) - 40;
                    bnTwo.Location = p;

                    bnTwo.Visible = true;
                    bnOk.Visible = true;
                    break;
            }
        }

        private void bnTwo_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}