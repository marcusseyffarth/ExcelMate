using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExcelMate
{
    public partial class RoundSelector : Form
    {
        private int intRound = 0;

        public RoundSelector(bool single)
        {
            InitializeComponent();

            cbOverwrite.SelectedIndex = 0;
            if (single)
            {
                cbOverwrite.Items.Add("Third run");
                cbOverwrite.Items.Add("Fourth run");
            }
        }

        private void bnOverWrite_Click(object sender, EventArgs e)
        {
            if (this.selectedRound != 0)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("Please select a value");
            }
        }

        private void bnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void cbOverwrite_SelectedIndexChanged(object sender, EventArgs e)
        {
            intRound = cbOverwrite.SelectedIndex;
        }

        public int selectedRound {
            get { return intRound; }
        }
    }
}