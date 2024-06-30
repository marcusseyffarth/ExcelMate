namespace ExcelMate
{
    using System.Drawing;
    using System.Windows.Forms;
    using System;
    using System.Diagnostics;

    partial class mExcelMate
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(mExcelMate));
            this.bnSaveLeft = new System.Windows.Forms.Button();
            this.tbLeftRaw = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.bnOpenWorkBook = new System.Windows.Forms.Button();
            this.tbFileName = new System.Windows.Forms.TextBox();
            this.cbWorkSheet = new System.Windows.Forms.ComboBox();
            this.cbRound = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tbRightReaction = new System.Windows.Forms.TextBox();
            this.cbRigthRider = new System.Windows.Forms.ComboBox();
            this.bnSaveRight = new System.Windows.Forms.Button();
            this.tbRightRaw = new System.Windows.Forms.TextBox();
            this.tbLeftReaction = new System.Windows.Forms.TextBox();
            this.cbLeftRider = new System.Windows.Forms.ComboBox();
            this.cbComPort = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.bnConnect = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cbRightColor = new System.Windows.Forms.ComboBox();
            this.cbLeftColor = new System.Windows.Forms.ComboBox();
            this.cbColors = new System.Windows.Forms.CheckBox();
            this.cbLog2File = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.cbSingleLanePort = new System.Windows.Forms.CheckBox();
            this.cbDiscardReactionTimes = new System.Windows.Forms.CheckBox();
            this.bnHelp = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cbLayOut = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.cbLeftCones = new System.Windows.Forms.ComboBox();
            this.cbRightCones = new System.Windows.Forms.ComboBox();
            this.process1 = new System.Diagnostics.Process();
            this.tbPrevData = new System.Windows.Forms.TextBox();
            this.cbPreviousData = new System.Windows.Forms.CheckBox();
            this.bnReset = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabSettings = new System.Windows.Forms.TabPage();
            this.tabRace = new System.Windows.Forms.TabPage();
            this.groupRight = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.groupLeft = new System.Windows.Forms.GroupBox();
            this.bnRefreshList = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabSettings.SuspendLayout();
            this.tabRace.SuspendLayout();
            this.groupRight.SuspendLayout();
            this.groupLeft.SuspendLayout();
            this.SuspendLayout();
            // 
            // bnSaveLeft
            // 
            this.bnSaveLeft.BackColor = System.Drawing.Color.WhiteSmoke;
            this.bnSaveLeft.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bnSaveLeft.ForeColor = System.Drawing.Color.Black;
            this.bnSaveLeft.Location = new System.Drawing.Point(656, 8);
            this.bnSaveLeft.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.bnSaveLeft.Name = "bnSaveLeft";
            this.bnSaveLeft.Size = new System.Drawing.Size(109, 43);
            this.bnSaveLeft.TabIndex = 10;
            this.bnSaveLeft.Text = "Save";
            this.bnSaveLeft.UseVisualStyleBackColor = false;
            this.bnSaveLeft.Click += new System.EventHandler(this.bnSaveLeft_Click);
            // 
            // tbLeftRaw
            // 
            this.tbLeftRaw.BackColor = System.Drawing.Color.White;
            this.tbLeftRaw.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbLeftRaw.ForeColor = System.Drawing.Color.Black;
            this.tbLeftRaw.Location = new System.Drawing.Point(451, 16);
            this.tbLeftRaw.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.tbLeftRaw.Name = "tbLeftRaw";
            this.tbLeftRaw.Size = new System.Drawing.Size(87, 31);
            this.tbLeftRaw.TabIndex = 3;
            this.tbLeftRaw.TabStop = false;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // bnOpenWorkBook
            // 
            this.bnOpenWorkBook.Location = new System.Drawing.Point(258, 34);
            this.bnOpenWorkBook.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.bnOpenWorkBook.Name = "bnOpenWorkBook";
            this.bnOpenWorkBook.Size = new System.Drawing.Size(100, 23);
            this.bnOpenWorkBook.TabIndex = 3;
            this.bnOpenWorkBook.TabStop = false;
            this.bnOpenWorkBook.Text = "Select Excel file";
            this.bnOpenWorkBook.UseVisualStyleBackColor = true;
            this.bnOpenWorkBook.Click += new System.EventHandler(this.bnOpenWorkBook_Click);
            // 
            // tbFileName
            // 
            this.tbFileName.BackColor = System.Drawing.Color.White;
            this.tbFileName.ForeColor = System.Drawing.Color.Black;
            this.tbFileName.Location = new System.Drawing.Point(8, 35);
            this.tbFileName.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.tbFileName.Name = "tbFileName";
            this.tbFileName.Size = new System.Drawing.Size(243, 20);
            this.tbFileName.TabIndex = 4;
            this.tbFileName.TabStop = false;
            // 
            // cbWorkSheet
            // 
            this.cbWorkSheet.BackColor = System.Drawing.Color.White;
            this.cbWorkSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbWorkSheet.ForeColor = System.Drawing.Color.Black;
            this.cbWorkSheet.FormattingEnabled = true;
            this.cbWorkSheet.Location = new System.Drawing.Point(134, 59);
            this.cbWorkSheet.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbWorkSheet.Name = "cbWorkSheet";
            this.cbWorkSheet.Size = new System.Drawing.Size(118, 21);
            this.cbWorkSheet.TabIndex = 5;
            this.cbWorkSheet.TabStop = false;
            this.cbWorkSheet.SelectedIndexChanged += new System.EventHandler(this.cbWorkSheet_SelectedIndexChanged);
            // 
            // cbRound
            // 
            this.cbRound.BackColor = System.Drawing.Color.White;
            this.cbRound.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbRound.Enabled = false;
            this.cbRound.ForeColor = System.Drawing.Color.Black;
            this.cbRound.FormattingEnabled = true;
            this.cbRound.Location = new System.Drawing.Point(134, 84);
            this.cbRound.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbRound.Name = "cbRound";
            this.cbRound.Size = new System.Drawing.Size(118, 21);
            this.cbRound.TabIndex = 6;
            this.cbRound.TabStop = false;
            this.cbRound.SelectedIndexChanged += new System.EventHandler(this.cbRound_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 62);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Please select worksheet";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 85);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Please select round";
            // 
            // tbRightReaction
            // 
            this.tbRightReaction.BackColor = System.Drawing.Color.White;
            this.tbRightReaction.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbRightReaction.ForeColor = System.Drawing.Color.Black;
            this.tbRightReaction.Location = new System.Drawing.Point(370, 16);
            this.tbRightReaction.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.tbRightReaction.Name = "tbRightReaction";
            this.tbRightReaction.Size = new System.Drawing.Size(75, 31);
            this.tbRightReaction.TabIndex = 7;
            this.tbRightReaction.TabStop = false;
            // 
            // cbRigthRider
            // 
            this.cbRigthRider.BackColor = System.Drawing.Color.White;
            this.cbRigthRider.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbRigthRider.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbRigthRider.ForeColor = System.Drawing.Color.Black;
            this.cbRigthRider.FormattingEnabled = true;
            this.cbRigthRider.Location = new System.Drawing.Point(6, 16);
            this.cbRigthRider.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbRigthRider.Name = "cbRigthRider";
            this.cbRigthRider.Size = new System.Drawing.Size(358, 32);
            this.cbRigthRider.TabIndex = 6;
            // 
            // bnSaveRight
            // 
            this.bnSaveRight.BackColor = System.Drawing.Color.WhiteSmoke;
            this.bnSaveRight.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bnSaveRight.ForeColor = System.Drawing.Color.Black;
            this.bnSaveRight.Location = new System.Drawing.Point(656, 52);
            this.bnSaveRight.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.bnSaveRight.Name = "bnSaveRight";
            this.bnSaveRight.Size = new System.Drawing.Size(109, 34);
            this.bnSaveRight.TabIndex = 10;
            this.bnSaveRight.Text = "Save Right";
            this.bnSaveRight.UseVisualStyleBackColor = false;
            this.bnSaveRight.Visible = false;
            // 
            // tbRightRaw
            // 
            this.tbRightRaw.BackColor = System.Drawing.Color.White;
            this.tbRightRaw.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbRightRaw.ForeColor = System.Drawing.Color.Black;
            this.tbRightRaw.Location = new System.Drawing.Point(451, 16);
            this.tbRightRaw.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.tbRightRaw.Name = "tbRightRaw";
            this.tbRightRaw.Size = new System.Drawing.Size(87, 31);
            this.tbRightRaw.TabIndex = 8;
            this.tbRightRaw.TabStop = false;
            // 
            // tbLeftReaction
            // 
            this.tbLeftReaction.BackColor = System.Drawing.Color.White;
            this.tbLeftReaction.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbLeftReaction.ForeColor = System.Drawing.Color.Black;
            this.tbLeftReaction.Location = new System.Drawing.Point(370, 16);
            this.tbLeftReaction.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.tbLeftReaction.Name = "tbLeftReaction";
            this.tbLeftReaction.Size = new System.Drawing.Size(75, 31);
            this.tbLeftReaction.TabIndex = 2;
            this.tbLeftReaction.TabStop = false;
            // 
            // cbLeftRider
            // 
            this.cbLeftRider.BackColor = System.Drawing.Color.White;
            this.cbLeftRider.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbLeftRider.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbLeftRider.ForeColor = System.Drawing.Color.Black;
            this.cbLeftRider.FormattingEnabled = true;
            this.cbLeftRider.Location = new System.Drawing.Point(6, 16);
            this.cbLeftRider.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbLeftRider.Name = "cbLeftRider";
            this.cbLeftRider.Size = new System.Drawing.Size(358, 32);
            this.cbLeftRider.TabIndex = 1;
            // 
            // cbComPort
            // 
            this.cbComPort.BackColor = System.Drawing.Color.White;
            this.cbComPort.ForeColor = System.Drawing.Color.Black;
            this.cbComPort.FormattingEnabled = true;
            this.cbComPort.Location = new System.Drawing.Point(168, 10);
            this.cbComPort.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbComPort.Name = "cbComPort";
            this.cbComPort.Size = new System.Drawing.Size(84, 21);
            this.cbComPort.TabIndex = 19;
            this.cbComPort.TabStop = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(7, 14);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(155, 13);
            this.label5.TabIndex = 20;
            this.label5.Text = "Select COM port for TrackMate";
            // 
            // bnConnect
            // 
            this.bnConnect.BackColor = System.Drawing.Color.WhiteSmoke;
            this.bnConnect.ForeColor = System.Drawing.Color.Black;
            this.bnConnect.Location = new System.Drawing.Point(258, 9);
            this.bnConnect.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.bnConnect.Name = "bnConnect";
            this.bnConnect.Size = new System.Drawing.Size(100, 23);
            this.bnConnect.TabIndex = 21;
            this.bnConnect.TabStop = false;
            this.bnConnect.Text = "Connect!";
            this.bnConnect.UseVisualStyleBackColor = false;
            this.bnConnect.Click += new System.EventHandler(this.bnConnect_Click_1);
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.White;
            this.groupBox1.Controls.Add(this.cbRightColor);
            this.groupBox1.Controls.Add(this.cbLeftColor);
            this.groupBox1.Controls.Add(this.cbColors);
            this.groupBox1.Controls.Add(this.cbLog2File);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.cbSingleLanePort);
            this.groupBox1.Controls.Add(this.cbDiscardReactionTimes);
            this.groupBox1.Controls.Add(this.bnHelp);
            this.groupBox1.ForeColor = System.Drawing.Color.Black;
            this.groupBox1.Location = new System.Drawing.Point(379, -2);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.groupBox1.Size = new System.Drawing.Size(506, 108);
            this.groupBox1.TabIndex = 22;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Advanced";
            // 
            // cbRightColor
            // 
            this.cbRightColor.FormattingEnabled = true;
            this.cbRightColor.Items.AddRange(new object[] {
            "Right",
            "White",
            "Red",
            "Green",
            "Orange",
            "Blue",
            "Yellow"});
            this.cbRightColor.Location = new System.Drawing.Point(224, 84);
            this.cbRightColor.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbRightColor.Name = "cbRightColor";
            this.cbRightColor.Size = new System.Drawing.Size(63, 21);
            this.cbRightColor.TabIndex = 47;
            // 
            // cbLeftColor
            // 
            this.cbLeftColor.FormattingEnabled = true;
            this.cbLeftColor.Items.AddRange(new object[] {
            "Left",
            "White",
            "Red",
            "Green",
            "Orange",
            "Blue",
            "Yellow"});
            this.cbLeftColor.Location = new System.Drawing.Point(151, 84);
            this.cbLeftColor.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbLeftColor.Name = "cbLeftColor";
            this.cbLeftColor.Size = new System.Drawing.Size(63, 21);
            this.cbLeftColor.TabIndex = 46;
            // 
            // cbColors
            // 
            this.cbColors.AutoSize = true;
            this.cbColors.Location = new System.Drawing.Point(6, 87);
            this.cbColors.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbColors.Name = "cbColors";
            this.cbColors.Size = new System.Drawing.Size(128, 17);
            this.cbColors.TabIndex = 45;
            this.cbColors.Text = "Use colorbased lanes";
            this.cbColors.UseVisualStyleBackColor = true;
            this.cbColors.CheckedChanged += new System.EventHandler(this.cbColors_CheckedChanged);
            // 
            // cbLog2File
            // 
            this.cbLog2File.AutoSize = true;
            this.cbLog2File.Checked = true;
            this.cbLog2File.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbLog2File.Location = new System.Drawing.Point(6, 62);
            this.cbLog2File.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbLog2File.Name = "cbLog2File";
            this.cbLog2File.Size = new System.Drawing.Size(144, 17);
            this.cbLog2File.TabIndex = 44;
            this.cbLog2File.Text = "Log recorded times to file";
            this.cbLog2File.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(151, 58);
            this.button1.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(62, 23);
            this.button1.TabIndex = 42;
            this.button1.TabStop = false;
            this.button1.Text = "Logfile";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.bnLogfile_Click);
            // 
            // cbSingleLanePort
            // 
            this.cbSingleLanePort.AutoSize = true;
            this.cbSingleLanePort.Location = new System.Drawing.Point(6, 38);
            this.cbSingleLanePort.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbSingleLanePort.Name = "cbSingleLanePort";
            this.cbSingleLanePort.Size = new System.Drawing.Size(282, 17);
            this.cbSingleLanePort.TabIndex = 43;
            this.cbSingleLanePort.Text = "Use Right lane port on Trackmate in single lane racing";
            this.cbSingleLanePort.UseVisualStyleBackColor = true;
            // 
            // cbDiscardReactionTimes
            // 
            this.cbDiscardReactionTimes.AutoSize = true;
            this.cbDiscardReactionTimes.Location = new System.Drawing.Point(6, 14);
            this.cbDiscardReactionTimes.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbDiscardReactionTimes.Name = "cbDiscardReactionTimes";
            this.cbDiscardReactionTimes.Size = new System.Drawing.Size(130, 17);
            this.cbDiscardReactionTimes.TabIndex = 42;
            this.cbDiscardReactionTimes.Text = "Discard reaction times";
            this.cbDiscardReactionTimes.UseVisualStyleBackColor = true;
            // 
            // bnHelp
            // 
            this.bnHelp.BackColor = System.Drawing.Color.WhiteSmoke;
            this.bnHelp.ForeColor = System.Drawing.Color.Black;
            this.bnHelp.Location = new System.Drawing.Point(443, 83);
            this.bnHelp.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.bnHelp.Name = "bnHelp";
            this.bnHelp.Size = new System.Drawing.Size(56, 23);
            this.bnHelp.TabIndex = 41;
            this.bnHelp.Text = "Help";
            this.bnHelp.UseVisualStyleBackColor = false;
            this.bnHelp.Click += new System.EventHandler(this.bnHelp_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.Color.White;
            this.groupBox2.Controls.Add(this.cbLayOut);
            this.groupBox2.Controls.Add(this.cbWorkSheet);
            this.groupBox2.Controls.Add(this.cbRound);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.bnOpenWorkBook);
            this.groupBox2.Controls.Add(this.cbComPort);
            this.groupBox2.Controls.Add(this.tbFileName);
            this.groupBox2.Controls.Add(this.bnConnect);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.ForeColor = System.Drawing.Color.Black;
            this.groupBox2.Location = new System.Drawing.Point(0, -2);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.groupBox2.Size = new System.Drawing.Size(373, 108);
            this.groupBox2.TabIndex = 23;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Basic settings";
            // 
            // cbLayOut
            // 
            this.cbLayOut.AutoSize = true;
            this.cbLayOut.Location = new System.Drawing.Point(258, 87);
            this.cbLayOut.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbLayOut.Name = "cbLayOut";
            this.cbLayOut.Size = new System.Drawing.Size(73, 17);
            this.cbLayOut.TabIndex = 41;
            this.cbLayOut.Text = "Old layout";
            this.cbLayOut.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.ForeColor = System.Drawing.Color.Black;
            this.label7.Location = new System.Drawing.Point(365, -1);
            this.label7.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(50, 13);
            this.label7.TabIndex = 24;
            this.label7.Text = "Reaction";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.ForeColor = System.Drawing.Color.Black;
            this.label8.Location = new System.Drawing.Point(446, -1);
            this.label8.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(51, 13);
            this.label8.TabIndex = 25;
            this.label8.Text = "Raw time";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.ForeColor = System.Drawing.Color.Black;
            this.label9.Location = new System.Drawing.Point(540, -1);
            this.label9.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(37, 13);
            this.label9.TabIndex = 26;
            this.label9.Text = "Cones";
            // 
            // cbLeftCones
            // 
            this.cbLeftCones.BackColor = System.Drawing.Color.White;
            this.cbLeftCones.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbLeftCones.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbLeftCones.ForeColor = System.Drawing.Color.Black;
            this.cbLeftCones.FormattingEnabled = true;
            this.cbLeftCones.Items.AddRange(new object[] {
            "0",
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15",
            "16",
            "17",
            "18",
            "19",
            "20",
            "21",
            "22",
            "23",
            "24",
            "25",
            "26",
            "27",
            "28",
            "29",
            "30",
            "DQ"});
            this.cbLeftCones.Location = new System.Drawing.Point(545, 14);
            this.cbLeftCones.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbLeftCones.Name = "cbLeftCones";
            this.cbLeftCones.Size = new System.Drawing.Size(70, 32);
            this.cbLeftCones.TabIndex = 4;
            this.cbLeftCones.SelectionChangeCommitted += new System.EventHandler(this.LeftConesChanged);
            // 
            // cbRightCones
            // 
            this.cbRightCones.BackColor = System.Drawing.Color.White;
            this.cbRightCones.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbRightCones.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbRightCones.ForeColor = System.Drawing.Color.Black;
            this.cbRightCones.FormattingEnabled = true;
            this.cbRightCones.Items.AddRange(new object[] {
            "0",
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15",
            "16",
            "17",
            "18",
            "19",
            "20",
            "21",
            "22",
            "23",
            "24",
            "25",
            "26",
            "27",
            "28",
            "29",
            "30",
            "DQ"});
            this.cbRightCones.Location = new System.Drawing.Point(544, 16);
            this.cbRightCones.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbRightCones.Name = "cbRightCones";
            this.cbRightCones.Size = new System.Drawing.Size(70, 32);
            this.cbRightCones.TabIndex = 9;
            this.cbRightCones.SelectionChangeCommitted += new System.EventHandler(this.RightConesChanged);
            // 
            // process1
            // 
            this.process1.StartInfo.Domain = "";
            this.process1.StartInfo.LoadUserProfile = false;
            this.process1.StartInfo.Password = null;
            this.process1.StartInfo.StandardErrorEncoding = null;
            this.process1.StartInfo.StandardOutputEncoding = null;
            this.process1.StartInfo.UserName = "";
            this.process1.SynchronizingObject = this;
            // 
            // tbPrevData
            // 
            this.tbPrevData.Font = new System.Drawing.Font("Verdana", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbPrevData.ForeColor = System.Drawing.Color.Black;
            this.tbPrevData.Location = new System.Drawing.Point(8, 105);
            this.tbPrevData.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.tbPrevData.Multiline = true;
            this.tbPrevData.Name = "tbPrevData";
            this.tbPrevData.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbPrevData.Size = new System.Drawing.Size(636, 101);
            this.tbPrevData.TabIndex = 29;
            this.tbPrevData.TabStop = false;
            this.tbPrevData.Visible = false;
            // 
            // cbPreviousData
            // 
            this.cbPreviousData.AutoSize = true;
            this.cbPreviousData.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.cbPreviousData.ForeColor = System.Drawing.Color.Black;
            this.cbPreviousData.Location = new System.Drawing.Point(656, 87);
            this.cbPreviousData.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cbPreviousData.Name = "cbPreviousData";
            this.cbPreviousData.Size = new System.Drawing.Size(94, 17);
            this.cbPreviousData.TabIndex = 31;
            this.cbPreviousData.TabStop = false;
            this.cbPreviousData.Text = "show previous";
            this.cbPreviousData.UseVisualStyleBackColor = true;
            this.cbPreviousData.CheckedChanged += new System.EventHandler(this.cbPreviousData_CheckedChanged);
            // 
            // bnReset
            // 
            this.bnReset.BackColor = System.Drawing.Color.WhiteSmoke;
            this.bnReset.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bnReset.ForeColor = System.Drawing.Color.Black;
            this.bnReset.Location = new System.Drawing.Point(770, 8);
            this.bnReset.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.bnReset.Name = "bnReset";
            this.bnReset.Size = new System.Drawing.Size(108, 43);
            this.bnReset.TabIndex = 11;
            this.bnReset.Text = "Reset/Start";
            this.bnReset.UseVisualStyleBackColor = false;
            this.bnReset.Click += new System.EventHandler(this.bnReset_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabSettings);
            this.tabControl1.Controls.Add(this.tabRace);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.Padding = new System.Drawing.Point(10, 3);
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(894, 136);
            this.tabControl1.TabIndex = 32;
            this.tabControl1.TabStop = false;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabSettings
            // 
            this.tabSettings.BackColor = System.Drawing.Color.White;
            this.tabSettings.Controls.Add(this.groupBox1);
            this.tabSettings.Controls.Add(this.groupBox2);
            this.tabSettings.Location = new System.Drawing.Point(4, 22);
            this.tabSettings.Margin = new System.Windows.Forms.Padding(0);
            this.tabSettings.Name = "tabSettings";
            this.tabSettings.Size = new System.Drawing.Size(886, 110);
            this.tabSettings.TabIndex = 0;
            this.tabSettings.Text = "Settings";
            // 
            // tabRace
            // 
            this.tabRace.BackColor = System.Drawing.Color.White;
            this.tabRace.Controls.Add(this.groupRight);
            this.tabRace.Controls.Add(this.groupLeft);
            this.tabRace.Controls.Add(this.bnRefreshList);
            this.tabRace.Controls.Add(this.bnReset);
            this.tabRace.Controls.Add(this.cbPreviousData);
            this.tabRace.Controls.Add(this.tbPrevData);
            this.tabRace.Controls.Add(this.bnSaveLeft);
            this.tabRace.Controls.Add(this.bnSaveRight);
            this.tabRace.Location = new System.Drawing.Point(4, 22);
            this.tabRace.Margin = new System.Windows.Forms.Padding(0);
            this.tabRace.Name = "tabRace";
            this.tabRace.Size = new System.Drawing.Size(886, 110);
            this.tabRace.TabIndex = 1;
            this.tabRace.Text = "Race";
            // 
            // groupRight
            // 
            this.groupRight.Controls.Add(this.cbRigthRider);
            this.groupRight.Controls.Add(this.tbRightReaction);
            this.groupRight.Controls.Add(this.tbRightRaw);
            this.groupRight.Controls.Add(this.label6);
            this.groupRight.Controls.Add(this.cbRightCones);
            this.groupRight.Controls.Add(this.label13);
            this.groupRight.Controls.Add(this.label14);
            this.groupRight.Location = new System.Drawing.Point(0, 52);
            this.groupRight.Margin = new System.Windows.Forms.Padding(0);
            this.groupRight.Name = "groupRight";
            this.groupRight.Padding = new System.Windows.Forms.Padding(0);
            this.groupRight.Size = new System.Drawing.Size(644, 53);
            this.groupRight.TabIndex = 39;
            this.groupRight.TabStop = false;
            this.groupRight.Text = "Right";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.ForeColor = System.Drawing.Color.Black;
            this.label6.Location = new System.Drawing.Point(541, -1);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(37, 13);
            this.label6.TabIndex = 34;
            this.label6.Text = "Cones";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.ForeColor = System.Drawing.Color.Black;
            this.label13.Location = new System.Drawing.Point(446, -1);
            this.label13.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(51, 13);
            this.label13.TabIndex = 33;
            this.label13.Text = "Raw time";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.ForeColor = System.Drawing.Color.Black;
            this.label14.Location = new System.Drawing.Point(366, -1);
            this.label14.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(50, 13);
            this.label14.TabIndex = 32;
            this.label14.Text = "Reaction";
            // 
            // groupLeft
            // 
            this.groupLeft.Controls.Add(this.cbLeftRider);
            this.groupLeft.Controls.Add(this.tbLeftRaw);
            this.groupLeft.Controls.Add(this.cbLeftCones);
            this.groupLeft.Controls.Add(this.tbLeftReaction);
            this.groupLeft.Controls.Add(this.label7);
            this.groupLeft.Controls.Add(this.label8);
            this.groupLeft.Controls.Add(this.label9);
            this.groupLeft.Location = new System.Drawing.Point(0, 0);
            this.groupLeft.Margin = new System.Windows.Forms.Padding(0);
            this.groupLeft.Name = "groupLeft";
            this.groupLeft.Padding = new System.Windows.Forms.Padding(0);
            this.groupLeft.Size = new System.Drawing.Size(644, 53);
            this.groupLeft.TabIndex = 38;
            this.groupLeft.TabStop = false;
            this.groupLeft.Text = "Left";
            // 
            // bnRefreshList
            // 
            this.bnRefreshList.Location = new System.Drawing.Point(770, 84);
            this.bnRefreshList.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.bnRefreshList.Name = "bnRefreshList";
            this.bnRefreshList.Size = new System.Drawing.Size(109, 19);
            this.bnRefreshList.TabIndex = 37;
            this.bnRefreshList.TabStop = false;
            this.bnRefreshList.Text = "Refresh racers";
            this.bnRefreshList.UseVisualStyleBackColor = true;
            this.bnRefreshList.Click += new System.EventHandler(this.bnRefreshList_Click);
            // 
            // mExcelMate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(894, 136);
            this.Controls.Add(this.tabControl1);
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "mExcelMate";
            this.Text = "Excelmate - Trackmate reader for skateboard slalom racing v6.0 - build 2024 06 30" +
    "";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ExcelMate_FormClosing);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabSettings.ResumeLayout(false);
            this.tabRace.ResumeLayout(false);
            this.tabRace.PerformLayout();
            this.groupRight.ResumeLayout(false);
            this.groupRight.PerformLayout();
            this.groupLeft.ResumeLayout(false);
            this.groupLeft.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion


        private System.Windows.Forms.CheckBox cbColors;

        private System.Windows.Forms.Button bnSaveLeft;
        private System.Windows.Forms.TextBox tbLeftRaw;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button bnOpenWorkBook;
        private System.Windows.Forms.TextBox tbFileName;
        private System.Windows.Forms.ComboBox cbWorkSheet;
        private System.Windows.Forms.ComboBox cbRound;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbLeftRider;
        private System.Windows.Forms.TextBox tbLeftReaction;
        private System.Windows.Forms.TextBox tbRightReaction;
        private System.Windows.Forms.ComboBox cbRigthRider;
        private System.Windows.Forms.Button bnSaveRight;
        private System.Windows.Forms.TextBox tbRightRaw;
        private System.Windows.Forms.ComboBox cbComPort;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button bnConnect;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cbLeftCones;
        private System.Windows.Forms.ComboBox cbRightCones;
        private System.Diagnostics.Process process1;
        private System.Windows.Forms.CheckBox cbPreviousData;
        private System.Windows.Forms.TextBox tbPrevData;
        private System.Windows.Forms.Button bnReset;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabSettings;
        private System.Windows.Forms.TabPage tabRace;
        private System.Windows.Forms.TextBox tbLiveId;
        private System.Windows.Forms.TextBox tbLiveEventId;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.CheckBox cbLayOut;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button bnCheckId;
        private System.Windows.Forms.CheckBox cbDiscardReactionTimes;
        private System.Windows.Forms.Button bnRefreshList;
        private System.Windows.Forms.CheckBox cbSingleLanePort;
        private System.Windows.Forms.ComboBox cbLeftColor;
        private System.Windows.Forms.ComboBox cbRightColor;
        private System.Windows.Forms.CheckBox cbLog2File;
        private System.Windows.Forms.CheckBox cbColor;
        private System.Windows.Forms.GroupBox groupRight;
        private System.Windows.Forms.GroupBox groupLeft;
        private System.Windows.Forms.Button button1;
        private Button bnHelp;
    }
}

