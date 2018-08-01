namespace ExcelMate
{
    using Mac.Excel9.Interop;
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Globalization;
    using System.IO;
    using System.IO.Ports;
    using System.Net;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Windows.Forms;

    public partial class mExcelMate : Form
    {
        private Thread readThread;
        private Thread writeThreadLeft;
        private Thread writeThreadRight;
        private ManualResetEvent m_EventStopThread;
        private ManualResetEvent m_EventResetTimer;
        private ManualResetEvent m_EventStopWriteToDisplayLeft;
        private ManualResetEvent m_EventStopWriteToDisplayRight;
        public TrackMateCallback mTrackMateCallback;
        private Mac.Excel9.Interop.Application objExcel = null;
        private Workbook theWorkbook = null;
        private static RaceTypes m_nRaceType = RaceTypes.SingleLane;
        private string m_strReportSiteSelection = "";
        private string m_strLogFile = "";
        private bool m_bolTrackmateVersion = false;
        private bool m_bolDisplayZeroReaction = false;
        private string m_strWorkBookName = "";
        private bool m_bolSendURL = true;
        private Color leftColor = Color.White;
        private Color rightColor = Color.Red;

        public mExcelMate()
        {
            InitializeComponent();
            try
            {
                this.objExcel = new ApplicationClass();
            }
            catch (Exception)
            {
                new MacMessageBox("The startup of this application failed, probably because you do not have Excel installed on the computer.", MessageBoxButtons.OK) { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                return;
            }
            this.mTrackMateCallback = new TrackMateCallback(this.AddTrackMateMessage);
            this.m_EventStopThread = new ManualResetEvent(false);
            this.m_EventResetTimer = new ManualResetEvent(false);
            this.m_EventStopWriteToDisplayLeft = new ManualResetEvent(false);
            this.m_EventStopWriteToDisplayRight = new ManualResetEvent(false);
            foreach (string str in SerialPort.GetPortNames())
            {
                this.cbComPort.Items.Add(str);
                this.cbDisplayPortLeft.Items.Add(str);
                this.cbDisplayPortRight.Items.Add(str);
            }
            this.cbComPort.SelectedIndex = -1;
            this.cbDisplayPortLeft.SelectedIndex = -1;
            this.cbDisplayPortRight.SelectedIndex = -1;
            this.cbDisplays.Checked = false;
            this.tbLeftRaw.Text = "";
            this.tbLeftReaction.Text = "";
            this.cbLeftCones.SelectedIndex = 0;
            this.cbLeftRider.SelectedIndex = -1;
            this.tbRightRaw.Text = "";
            this.tbRightReaction.Text = "";
            this.cbRightCones.SelectedIndex = 0;
            this.cbRigthRider.SelectedIndex = -1;
            this.m_strLogFile = "RaceLog_" + DateTime.Today.ToString("yyyyMMdd") + ".log";
            this.cbRightColor.Enabled = false;
            this.cbLeftColor.Enabled = false;
            this.cbRightColor.SelectedIndex = 0;
            this.cbLeftColor.SelectedIndex = 0;
        }

        private void AddTrackMateMessage(string message, int lane, int type)
        {
            string str;
            string text;
            string str3;
            string str4;
            if (type == 0)
            {
                if (!this.cbDiscardReactionTimes.Checked)
                {
                    if ((lane == 1) || ((m_nRaceType == RaceTypes.SingleLane) && this.cbSingleLanePort.Checked))
                    {
                        this.tbLeftReaction.Text = message;
                        if ((this.cbDisplays.Checked && this.m_bolTrackmateVersion) && (this.writeThreadLeft == null))
                        {
                            this.StartDisplayThreadLeft();
                        }
                    }
                    else
                    {
                        this.tbRightReaction.Text = message;
                        if ((this.cbDisplays.Checked && this.m_bolTrackmateVersion) && (this.writeThreadRight == null))
                        {
                            this.StartDisplayThreadRight();
                        }
                    }
                }
            }
            else if ((lane == 1) || ((m_nRaceType == RaceTypes.SingleLane) && this.cbSingleLanePort.Checked))
            {
                this.tbLeftRaw.Text = message;
            }
            else
            {
                this.tbRightRaw.Text = message;
            }
            this.WriteToLogfile(message + " lane: " + lane.ToString());
            string selectedItem = "";
            if (lane == 1)
            {
                str = this.cbLeftRider.SelectedItem.ToString().Trim();
                text = this.tbLeftReaction.Text;
                str3 = this.tbLeftRaw.Text.Replace(".", ",");
                str4 = this.cbLeftCones.SelectedItem.ToString();
                selectedItem = (string)this.cbLeftColor.SelectedItem;
            }
            else
            {
                str = this.cbRigthRider.SelectedItem.ToString().Trim();
                text = this.tbRightReaction.Text;
                str3 = this.tbRightRaw.Text.Replace(".", ",");
                str4 = this.cbRightCones.SelectedItem.ToString();
                selectedItem = (string)this.cbRightColor.SelectedItem;
            }
            if (m_nRaceType == RaceTypes.SingleLane)
            {
                selectedItem = "R";
            }
            this.submitRunToWeb(str, text, str3, str4, selectedItem, false);
        }

        private void bnCheckId_Click(object sender, EventArgs e)
        {
            string fileName = "http://www.slalomskateboarder.com/ISSA/Racing_log/ISSA_race_log.pdf";
            Process.Start(fileName);
        }

        private void bnConnect_Click_1(object sender, EventArgs e)
        {
            if (this.bnConnect.Text == "Connect!")
            {
                this.startThread();
            }
            else
            {
                this.StopThread();
                this.bnConnect.Text = "Connect!";
            }
        }

        private void bnLogfile_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.FileName = "*.log";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.m_strLogFile = this.openFileDialog1.FileName;
            }
        }

        private void bnOpenWorkBook_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            m_nRaceType = RaceTypes.NotSet;
            this.openFileDialog1.FileName = "*.xls";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.tbFileName.Text = this.openFileDialog1.FileName;
                try
                {
                    CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
                    Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                    if (this.theWorkbook != null)
                    {
                        this.theWorkbook.Save();
                        this.theWorkbook.Close(true, this.m_strWorkBookName, false);
                        this.objExcel.Visible = false;
                        this.objExcel = null;
                        Thread.Sleep(0x3e8);
                        this.objExcel = new ApplicationClass();
                    }
                    this.theWorkbook = this.objExcel.Workbooks.Open(this.tbFileName.Text, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true);
                    Thread.CurrentThread.CurrentCulture = currentCulture;
                }
                catch (Exception exception)
                {
                    new MacMessageBox("I guess you were not able to open the workbook. Please make sure it is available and not readonly. " + exception.Message) { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                    this.tbFileName.Text = "";
                    return;
                }
                Sheets worksheets = this.theWorkbook.Worksheets;
                this.cbWorkSheet.Items.Clear();
                for (int i = 1; i <= worksheets.Count; i++)
                {
                    _Worksheet worksheet = (_Worksheet)worksheets.get_Item(i);
                    this.cbWorkSheet.Items.Add(worksheet.Name);
                }
                this.Text = this.tbFileName.Text;
                this.m_strWorkBookName = this.tbFileName.Text;
            }
            this.Cursor = Cursors.Default;
        }

        private void bnRefreshList_Click(object sender, EventArgs e)
        {
            if ((m_nRaceType == RaceTypes.SingleLane) || (m_nRaceType == RaceTypes.Qualification))
            {
                _Worksheet wS = null;
                CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                wS = (_Worksheet)this.theWorkbook.Worksheets.get_Item(this.cbWorkSheet.SelectedIndex + 1);
                wS.Activate();
                this.FillRiderDropDownsSingle(wS);
                Thread.CurrentThread.CurrentCulture = currentCulture;
            }
        }

        private void bnReset_Click(object sender, EventArgs e)
        {
            DialogResult no = DialogResult.No;
            if ((this.tbLeftRaw.Text.Length != 0) || (this.tbRightRaw.Text.Length != 0))
            {
                no = MessageBox.Show("It seems like you haven't saved the latest data. Would you like to save first?", "Forgot to save?", MessageBoxButtons.YesNo);
            }
            if (no == DialogResult.No)
            {
                if ((this.readThread != null) && this.readThread.IsAlive)
                {
                    this.m_EventResetTimer.Set();
                    Thread.Sleep(0x3e8);
                    this.m_EventResetTimer.Reset();
                    this.tbLeftRaw.Text = "";
                    this.tbRightRaw.Text = "";
                    this.tbLeftReaction.Text = "";
                    this.tbRightReaction.Text = "";
                    this.m_bolSendURL = true;
                    if (this.cbDisplays.Checked && !this.m_bolTrackmateVersion)
                    {
                        if (this.writeThreadLeft == null)
                        {
                            this.StartDisplayThreadLeft();
                        }
                        if (this.writeThreadRight == null)
                        {
                            this.StartDisplayThreadRight();
                        }
                    }
                    if (this.cbLiveReport.Checked)
                    {
                        if (this.tbEventId.Text.Length == 0)
                        {
                            new MacMessageBox("You need to fill in the Live Reporter Id in order to submit times to the web. Get your own reporter Id, send an email to marcus@ettsexett.com.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                        }
                        else if (this.tbReporterId.Text.Length == 0)
                        {
                            new MacMessageBox("You need to fill in the EventId in order to submit times to the web. To get the event Id, send an email to marcus@ettsexett.com.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                        }
                        else
                        {
                            this.submitRunToWeb("", "", "", "", "", true);
                        }
                    }
                }
                else
                {
                    new MacMessageBox("I don't think you got a connection to the Trackmate.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                }
            }
        }

        private void bnSaveLeft_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            switch (m_nRaceType)
            {
                case RaceTypes.SingleLane:
                    this.saveSingleLaneRace();
                    this.m_bolSendURL = false;
                    break;

                case RaceTypes.Qualification:
                    {
                        bool flag = false;
                        flag = this.saveDualQualification(Lane.Left);
                        if (this.saveDualQualification(Lane.Right))
                        {
                            this.SaveToExcel();
                            this.m_bolSendURL = false;
                        }
                        break;
                    }
                case RaceTypes.Elimination:
                    {
                        bool flag3 = false;
                        flag3 = this.saveDualLaneRace(Lane.Left);
                        if (this.saveDualLaneRace(Lane.Right))
                        {
                            this.SaveToExcel();
                            this.m_bolSendURL = false;
                        }
                        break;
                    }
            }
            this.Cursor = Cursors.Default;
        }

        private void bnTest_Click(object sender, EventArgs e)
        {
            if (((this.tabControl1.SelectedIndex != 0) && this.cbDisplays.Checked) && (((this.cbComPort.Text == this.cbDisplayPortLeft.Text) || (this.cbComPort.Text == this.cbDisplayPortRight.Text)) || (this.cbDisplayPortLeft.Text == this.cbDisplayPortRight.Text)))
            {
                new MacMessageBox("You need to select different comports for each connection if you're using external displays.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
            }
            else
            {
                new MacMessageBox("Connected!") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
            }
        }

        private void cbColors_CheckedChanged(object sender, EventArgs e)
        {
            if (this.cbColors.Checked)
            {
                this.cbLeftColor.Enabled = true;
                this.cbRightColor.Enabled = true;
            }
            else
            {
                this.cbLeftColor.Enabled = false;
                this.cbRightColor.Enabled = false;
                this.cbLeftColor.SelectedIndex = 0;
                this.cbRightColor.SelectedIndex = 0;
            }
        }

        private void cbDisplays_CheckedChanged(object sender, EventArgs e)
        {
            if (this.cbDisplays.Checked)
            {
                if (MessageBox.Show("If you have Trackmate v6.5 or higher click 'Yes' otherwise click 'No' and the times on the displays will start rolling at the fourth beep.", "Trackmate version", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    this.m_bolTrackmateVersion = true;
                }
                else
                {
                    this.m_bolTrackmateVersion = false;
                }
                if (MessageBox.Show("If this race uses individual starts (no reaction time) click yes, otherwise click no", "Use reaction", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    this.m_bolDisplayZeroReaction = true;
                }
                else
                {
                    this.m_bolDisplayZeroReaction = false;
                }
                this.label10.Visible = true;
                this.label11.Visible = true;
                this.label12.Visible = true;
                this.cbDisplayPortLeft.Visible = true;
                this.cbDisplayPortRight.Visible = true;
                this.bnConnectDisplays.Visible = true;
            }
            else
            {
                this.label10.Visible = false;
                this.label11.Visible = false;
                this.label12.Visible = false;
                this.cbDisplayPortLeft.Visible = false;
                this.cbDisplayPortRight.Visible = false;
                this.bnConnectDisplays.Visible = false;
            }
        }

        private void cbLiveReport_CheckedChanged(object sender, EventArgs e)
        {
            if (this.cbLiveReport.Checked)
            {
                this.label15.Visible = true;
                this.tbEventId.Visible = true;
                this.tbEventId.Enabled = true;
                this.label16.Visible = true;
                this.tbReporterId.Visible = true;
                this.tbReporterId.Enabled = true;
            }
            else
            {
                this.label15.Visible = false;
                this.tbEventId.Visible = false;
                this.tbEventId.Enabled = false;
                this.label16.Visible = false;
                this.tbReporterId.Visible = false;
                this.tbReporterId.Enabled = false;
            }
        }

        private void cbPreviousData_CheckedChanged(object sender, EventArgs e)
        {
            if (m_nRaceType == RaceTypes.SingleLane)
            {
                if (this.cbPreviousData.Checked)
                {
                    this.tbPrevData.Visible = true;
                    base.Height = 270;
                }
                else
                {
                    this.tbPrevData.Visible = false;
                    base.Height = 0x83;
                }
            }
            else if (this.cbPreviousData.Checked)
            {
                this.tbPrevData.Visible = true;
                base.Height = 270;
            }
            else
            {
                this.tbPrevData.Visible = false;
                base.Height = 0xa1;
            }
        }

        private void cbRound_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.cbRound.SelectedIndex != -1)
            {
                CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                this.cbLeftRider.Items.Clear();
                this.cbRigthRider.Items.Clear();
                _Worksheet worksheet = null;
                try
                {
                    worksheet = (_Worksheet)this.theWorkbook.Worksheets.get_Item(this.cbWorkSheet.SelectedIndex + 1);
                }
                catch (Exception exception)
                {
                    new MacMessageBox("Oops! We were not able to read the excelsheet that you previously selected. Did you kill it? Please re-open the spread sheet. " + exception.Message) { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                }
                Array values = (Array)worksheet.get_Range("B1", "B99").Cells.Value2;
                string[] strArray = this.ConvertToStringArray(values);
                string str = "";
                int num = 1;
                bool flag2 = false;
                string str2 = this.cbRound.SelectedItem.ToString();
                string str3 = "";
                foreach (string str4 in strArray)
                {
                    num++;
                    if (str2 == "Final & consi")
                    {
                        if (str4.ToLower().Contains("cons"))
                        {
                            str = "B" + num.ToString();
                            str3 = " - (consi)";
                        }
                    }
                    else if (str4.IndexOf(str2) != -1)
                    {
                        str = "B" + num.ToString();
                    }
                    if (str != "")
                    {
                        string str5 = "B" + num.ToString();
                        Range range2 = worksheet.get_Range(str5, str5);
                        if (range2.Value2 == null)
                        {
                            if ((str2 == "Final & consi") && !flag2)
                            {
                                flag2 = true;
                                str5 = "B" + ((num + 3)).ToString();
                                range2 = worksheet.get_Range(str5, str5);
                                str3 = " - (final)";
                            }
                            else
                            {
                                break;
                            }
                        }
                        string str6 = range2.Cells.Value2.ToString();
                        this.cbLeftRider.Items.Add(str6 + str3);
                        this.cbRigthRider.Items.Add(str6 + str3);
                    }
                }
                Thread.CurrentThread.CurrentCulture = currentCulture;
                this.tabControl1.SelectedIndex = 1;
            }
        }

        private void cbWorkSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            _Worksheet wS = null;
            try
            {
                this.bnRefreshList.Visible = true;
                CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                wS = (_Worksheet)this.theWorkbook.Worksheets.get_Item(this.cbWorkSheet.SelectedIndex + 1);
                wS.Activate();
                this.objExcel.Visible = true;
                base.TopMost = true;
                base.TopMost = false;
                this.cbRound.Items.Clear();
                if (wS.Name.ToLower().IndexOf("qual") != -1)
                {
                    new MacMessageBox("This race will be treated as qualification in head 2 head format (since the workbook contains 'qual').") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                    m_nRaceType = RaceTypes.Qualification;
                }
                else if (wS.Name.ToLower().IndexOf("elim") != -1)
                {
                    new MacMessageBox("This race will be treated as eliminatin in head 2 head format (since the workbook contains 'elim').") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                    m_nRaceType = RaceTypes.Elimination;
                }
                else
                {
                    new MacMessageBox("This race will be treated as single lane (since the workbook does not contain 'qual' or 'elim').") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                    m_nRaceType = RaceTypes.SingleLane;
                }
                if ((m_nRaceType == RaceTypes.SingleLane) || (m_nRaceType == RaceTypes.Qualification))
                {
                    this.FillRiderDropDownsSingle(wS);
                }
                else
                {
                    this.cbRound.Enabled = true;
                    Array values = (Array)wS.get_Range("B1", "B99").Cells.Value2;
                    string[] strArray = this.ConvertToStringArray(values);
                    foreach (string str in strArray)
                    {
                        if (str.ToLower().IndexOf("round") != -1)
                        {
                            if (str.ToLower().IndexOf("final") != -1)
                            {
                                this.cbRound.Items.Add("Final & consi");
                            }
                            else if (str.ToLower().IndexOf("cons") == -1)
                            {
                                this.cbRound.Items.Add(str);
                            }
                        }
                    }
                    if (this.cbRound.Items.Count == 0)
                    {
                        new MacMessageBox("No racers was found in the Excel sheet. If this is single lane or qualification remember that \r\nthe name of the workbook must contain 'qual' in order for the program to find the racers.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                    }
                }
                this.colorWorkSheet();
                Thread.CurrentThread.CurrentCulture = currentCulture;
            }
            catch (Exception exception)
            {
                new MacMessageBox("Oops! We were not able to read the excelsheet that you previously selected. Did you kill it? Please re-open the spread sheet. " + exception.Message) { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
            }
        }

        private void colorWorkSheet()
        {
            _Worksheet worksheet = null;
            worksheet = this.getWorkSheet();
            Range range = null;
            if (this.cbColors.Checked)
            {
                Color white = Color.White;
                Color red = Color.Red;
                if (((string)this.cbLeftColor.SelectedItem) == "White")
                {
                    white = Color.White;
                }
                if (((string)this.cbLeftColor.SelectedItem) == "Red")
                {
                    white = Color.Red;
                }
                if (((string)this.cbLeftColor.SelectedItem) == "Green")
                {
                    white = Color.LightGreen;
                }
                if (((string)this.cbLeftColor.SelectedItem) == "Orange")
                {
                    white = Color.Orange;
                }
                if (((string)this.cbLeftColor.SelectedItem) == "Blue")
                {
                    white = Color.LightBlue;
                }
                if (((string)this.cbLeftColor.SelectedItem) == "Yellow")
                {
                    white = Color.Yellow;
                }
                if (((string)this.cbRightColor.SelectedItem) == "White")
                {
                    red = Color.White;
                }
                if (((string)this.cbRightColor.SelectedItem) == "Red")
                {
                    red = Color.Red;
                }
                if (((string)this.cbRightColor.SelectedItem) == "Green")
                {
                    red = Color.LightGreen;
                }
                if (((string)this.cbRightColor.SelectedItem) == "Orange")
                {
                    red = Color.Orange;
                }
                if (((string)this.cbRightColor.SelectedItem) == "Blue")
                {
                    red = Color.LightBlue;
                }
                if (((string)this.cbRightColor.SelectedItem) == "Yellow")
                {
                    red = Color.Yellow;
                }
                for (int i = 12; i < 0x70; i++)
                {
                    if (((string)this.cbWorkSheet.SelectedItem).ToLower().Contains("qual"))
                    {
                        range = worksheet.get_Range("D" + i.ToString(), "I" + i.ToString());
                    }
                    else
                    {
                        range = worksheet.get_Range("C" + i.ToString(), "H" + i.ToString());
                    }
                    this.setBGColor(range, i, white, red);
                    range = worksheet.get_Range("K" + i.ToString(), "P" + i.ToString());
                    this.setBGColor(range, i + 1, white, red);
                }
                worksheet = null;
            }
        }

        private string[] ConvertToStringArray(Array values)
        {
            string[] strArray = new string[values.Length];
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(i, 1) == null)
                {
                    strArray[i - 1] = "";
                }
                else
                {
                    strArray[i - 1] = values.GetValue(i, 1).ToString();
                }
            }
            return strArray;
        }


        private void ExcelMate_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.onExit(null, null);
        }

        private void FillRiderDropDownsSingle(_Worksheet WS)
        {
            this.cbRound.Enabled = false;
            this.cbLeftRider.Items.Clear();
            this.cbRigthRider.Items.Clear();
            Array values = (Array)WS.get_Range("C1", "C99").Cells.Value2;
            string[] strArray = this.ConvertToStringArray(values);
            string str = "";
            int num = 2;
            foreach (string str2 in strArray)
            {
                if (str2.ToLower().IndexOf("name") != -1)
                {
                    num++;
                    str = "C" + num.ToString();
                }
                else
                {
                    num++;
                }
                if (str != "")
                {
                    string str3 = "C" + num.ToString();
                    string str4 = "B" + num.ToString();
                    string str5 = "";
                    Range range2 = WS.get_Range(str3, str3);
                    Range range3 = WS.get_Range(str4, str4);
                    if (range2.Value2 == null)
                    {
                        break;
                    }
                    if (range3.Value2 != null)
                    {
                        str5 = range3.Cells.Value2.ToString() + " ";
                    }
                    string item = str5 + range2.Cells.Value2.ToString();
                    this.cbLeftRider.Items.Add(item);
                    this.cbRigthRider.Items.Add(item);
                }
            }
            this.tabControl1.SelectedIndex = 1;
        }

        private int getClassId()
        {
            if (this.cbWorkSheet.SelectedItem.ToString().ToLower().Contains("am"))
            {
                return 1;
            }
            if (this.cbWorkSheet.SelectedItem.ToString().ToLower().Contains("pro"))
            {
                return 2;
            }
            if (this.cbWorkSheet.SelectedItem.ToString().ToLower().Contains("wo"))
            {
                return 3;
            }
            if (this.cbWorkSheet.SelectedItem.ToString().ToLower().Contains("jr") || this.cbWorkSheet.SelectedItem.ToString().ToLower().Contains("junior"))
            {
                return 4;
            }
            if (this.cbWorkSheet.SelectedItem.ToString().ToLower().Contains("mas"))
            {
                return 6;
            }
            return 5;
        }

        private Color GetColor(string name)
        {
            switch (name)
            {
                case "Red":
                    return Color.Red;

                case "Green":
                    return Color.LightGreen;

                case "Orange":
                    return Color.Orange;

                case "Blue":
                    return Color.LightBlue;

                case "Yellow":
                    return Color.Yellow;
            }
            return Color.White;
        }

        private int getRoundId()
        {
            if (this.cbRound.SelectedItem != null)
            {
                if (this.cbRound.SelectedItem.ToString().Contains("Final"))
                {
                    return 6;
                }
                if (this.cbRound.SelectedItem.ToString().Contains("4"))
                {
                    return 5;
                }
                if (this.cbRound.SelectedItem.ToString().Contains("8"))
                {
                    return 4;
                }
                if (this.cbRound.SelectedItem.ToString().Contains("16"))
                {
                    return 3;
                }
                if (this.cbRound.SelectedItem.ToString().Contains("32"))
                {
                    return 2;
                }
                if (this.cbRound.SelectedItem.ToString().Contains("64"))
                {
                    return 1;
                }
            }
            return 0;
        }

        private string GetTimeString(string tid)
        {
            string str = Convert.ToString(DateTime.Now.Hour);
            string str2 = Convert.ToString(DateTime.Now.Minute);
            string str3 = Convert.ToString(DateTime.Now.Second + tid.Substring(0, tid.IndexOf(".")));
            if (str.Length == 1)
            {
                str = "0" + str;
            }
            if (str2.Length == 1)
            {
                str2 = "0" + str2;
            }
            if (str3.Length == 1)
            {
                str3 = "0" + str3;
            }
            string[] textArray1 = new string[] { str, ":", str2, ":", str3 };
            return string.Concat(textArray1);
        }

        private _Worksheet getWorkSheet()
        {
            _Worksheet worksheet = null;
            try
            {
                worksheet = (_Worksheet)this.theWorkbook.Worksheets.get_Item(this.cbWorkSheet.SelectedIndex + 1);
            }
            catch (Exception exception)
            {
                new MacMessageBox("Oops! We were not able to read the excelsheet that you previously selected. Did you kill it? Please re-open the spread sheet. " + exception.Message) { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
            }
            return worksheet;
        }


        private bool isRawTimeEntered(Lane lane, bool showMsg)
        {
            if (lane == Lane.Left)
            {
                if (this.tbLeftRaw.Text.Length != 0)
                {
                    return true;
                }
            }
            else if (this.tbRightRaw.Text.Length != 0)
            {
                return true;
            }
            if (showMsg)
            {
                new MacMessageBox("There is no raw time to save.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
            }
            return false;
        }

        private bool isReactionTimeEntered(Lane lane, bool showMsg)
        {
            if (lane == Lane.Left)
            {
                if (this.tbLeftReaction.Text.Length != 0)
                {
                    return true;
                }
            }
            else if (this.tbRightReaction.Text.Length != 0)
            {
                return true;
            }
            if (showMsg)
            {
                new MacMessageBox("There is no reaction time to save.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
            }
            return false;
        }

        private bool isRiderSelected(Lane lane, bool showMsg)
        {
            if (lane == Lane.Left)
            {
                if (this.cbLeftRider.SelectedIndex != -1)
                {
                    return true;
                }
            }
            else if (this.cbRigthRider.SelectedIndex != -1)
            {
                return true;
            }
            if (showMsg)
            {
                new MacMessageBox("Please select a rider to save the time to.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
            }
            return false;
        }

        private void LeftConesChanged(object sender, EventArgs e)
        {
            if (this.cbLeftCones.SelectedItem.ToString() == "DQ")
            {
                MacMessageBox box = new MacMessageBox("Should this race be a DQ?", MessageBoxButtons.OKCancel)
                {
                    StartPosition = FormStartPosition.CenterParent
                };
                box.ShowDialog();
                if (box.DialogResult == DialogResult.OK)
                {
                    this.cbLeftCones.SelectedIndex = 0;
                    this.tbLeftRaw.Text = "DQ";
                }
            }
        }

        private void onExit(object sender, FormClosingEventArgs e)
        {
            this.StopThread();
            this.StopDisplayThreadLeft();
            this.StopDisplayThreadRight();
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            try
            {
                if (this.tbFileName.Text.Length != 0)
                {
                    this.theWorkbook.Save();
                    this.theWorkbook.Close(true, this.tbFileName.Text, false);
                }
            }
            catch (Exception exception)
            {
                new MacMessageBox("Oops! We couldn't close the Excel spread sheet properly. Perhaps you killed it already? " + exception.Message) { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
            }
            this.objExcel = null;
            Thread.CurrentThread.CurrentCulture = currentCulture;
        }

        private void RightConesChanged(object sender, EventArgs e)
        {
            if (this.cbRightCones.SelectedItem.ToString() == "DQ")
            {
                MacMessageBox box = new MacMessageBox("Should this race be a DQ?", MessageBoxButtons.OKCancel)
                {
                    StartPosition = FormStartPosition.CenterParent
                };
                box.ShowDialog();
                if (box.DialogResult == DialogResult.OK)
                {
                    this.cbRightCones.SelectedIndex = 0;
                    this.tbRightRaw.Text = "DQ";
                }
            }
        }

        private void RunDisplayThreadLeft()
        {
            DisplayWriter writer = new DisplayWriter(this, this.m_EventStopWriteToDisplayLeft);
            try
            {
                writer.openComPort(this.cbDisplayPortLeft.Text);
            }
            catch (Exception exception)
            {
                new MacMessageBox("Failed to open the com port for the left display, please close the program and start over. " + exception.Message) { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                return;
            }
            if (this.m_bolDisplayZeroReaction)
            {
                writer.WriteReactionToDisplay("0.000");
            }
            else
            {
                writer.WriteReactionToDisplay(this.tbLeftReaction.Text);
            }
        }

        private void RunDisplayThreadRight()
        {
            DisplayWriterRight right = new DisplayWriterRight(this, this.m_EventStopWriteToDisplayRight);
            try
            {
                right.openComPort(this.cbDisplayPortRight.Text);
            }
            catch (Exception exception)
            {
                new MacMessageBox("Failed to open the com port for the left display, please close the program and start over. " + exception.Message) { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                return;
            }
            if (this.m_bolDisplayZeroReaction)
            {
                right.WriteReactionToDisplay("0.000");
            }
            else
            {
                right.WriteReactionToDisplay(this.tbRightReaction.Text);
            }
        }

        private void RunFinalTimeDisplayThreadLeft()
        {
            DisplayWriter writer = new DisplayWriter(this, this.m_EventStopWriteToDisplayLeft);
            writer.openComPort(this.cbDisplayPortLeft.Text);
            writer.WriteFinalTimeToDisplay(this.tbLeftRaw.Text);
        }

        private void RunFinalTimeDisplayThreadRight()
        {
            DisplayWriterRight right = new DisplayWriterRight(this, this.m_EventStopWriteToDisplayRight);
            right.openComPort(this.cbDisplayPortRight.Text);
            right.WriteFinalTimeToDisplay(this.tbRightRaw.Text);
        }

        private void RunThread()
        {
            TrackMateReader reader = new TrackMateReader(this, this.m_EventStopThread, this.m_EventResetTimer);
            try
            {
                reader.openComPort(this.cbComPort.Text);
            }
            catch (Exception exception)
            {
                new MacMessageBox("Failed to open the com port for trackmate, please close the program and start over. " + exception.Message) { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                return;
            }
            MacMessageBox box = new MacMessageBox("Connected!")
            {
                StartPosition = FormStartPosition.Manual
            };
            System.Drawing.Point location = base.Location;
            location.X = (location.X + (base.Width / 2)) - (box.Width / 2);
            location.Y = (location.Y + (base.Height / 2)) - (box.Height / 2);
            box.Location = location;
            box.ShowDialog();
            reader.Read();
        }

        private void SaveAndReset(Lane lane)
        {
            string str;
            string text;
            string str3;
            string str4;
            string selectedItem;
            if (lane == Lane.Left)
            {
                str = this.cbLeftRider.SelectedItem.ToString().Trim();
                text = this.tbLeftReaction.Text;
                str3 = this.tbLeftRaw.Text.Replace(".", ",");
                str4 = this.cbLeftCones.SelectedItem.ToString();
                selectedItem = (string)this.cbLeftColor.SelectedItem;
            }
            else
            {
                str = this.cbRigthRider.SelectedItem.ToString().Trim();
                text = this.tbRightReaction.Text;
                str3 = this.tbRightRaw.Text.Replace(".", ",");
                str4 = this.cbRightCones.SelectedItem.ToString();
                selectedItem = (string)this.cbRightColor.SelectedItem;
            }
            if (m_nRaceType == RaceTypes.SingleLane)
            {
                selectedItem = "R";
            }
            if (this.cbLiveReport.Checked)
            {
                if (this.tbEventId.Text.Length == 0)
                {
                    new MacMessageBox("You need to fill in the Live Reporter Id in order to submit times to the web. Get your own reporter Id, send an email to marcus@ettsexett.com.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                    return;
                }
                if (this.tbReporterId.Text.Length == 0)
                {
                    new MacMessageBox("You need to fill in the EventId in order to submit times to the web. To get the event Id, send an email to marcus@ettsexett.com.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                    return;
                }
                this.submitRunToWeb(str, text, str3, str4, selectedItem, false);
            }
            string[] textArray1 = new string[] { str, " - ", text, ", ", str3, " + ", str4, Environment.NewLine, this.tbPrevData.Text };
            this.tbPrevData.Text = string.Concat(textArray1);
            if (lane == Lane.Left)
            {
                this.tbLeftReaction.Text = "";
                this.tbLeftReaction.Refresh();
                this.tbLeftRaw.Text = "";
                this.tbLeftRaw.Refresh();
                this.cbLeftCones.SelectedIndex = 0;
                if (m_nRaceType == RaceTypes.SingleLane)
                {
                    if (this.cbLeftRider.Items.Count > (this.cbLeftRider.SelectedIndex + 1))
                    {
                        this.cbLeftRider.SelectedIndex++;
                    }
                    else
                    {
                        this.cbLeftRider.SelectedIndex = -1;
                    }
                }
                else if (this.cbLeftRider.Items.Count > (this.cbLeftRider.SelectedIndex + 2))
                {
                    this.cbLeftRider.SelectedIndex += 2;
                }
                else
                {
                    this.cbLeftRider.SelectedIndex = -1;
                }
            }
            else
            {
                this.tbRightReaction.Text = "";
                this.tbRightReaction.Refresh();
                this.tbRightRaw.Text = "";
                this.tbRightRaw.Refresh();
                this.cbRightCones.SelectedIndex = 0;
                if (m_nRaceType == RaceTypes.SingleLane)
                {
                    if (this.cbRigthRider.Items.Count > (this.cbRigthRider.SelectedIndex + 1))
                    {
                        this.cbRigthRider.SelectedIndex++;
                    }
                    else
                    {
                        this.cbRigthRider.SelectedIndex = -1;
                    }
                }
                else if (this.cbRigthRider.Items.Count > (this.cbRigthRider.SelectedIndex + 2))
                {
                    this.cbRigthRider.SelectedIndex += 2;
                }
                else
                {
                    this.cbRigthRider.SelectedIndex = -1;
                }
            }
        }

        private bool saveDualLaneRace(Lane lane)
        {
            if ((!this.isRiderSelected(lane, false) && !this.isRawTimeEntered(lane, false)) && !this.isReactionTimeEntered(lane, false))
            {
                return false;
            }
            if (!this.cbDiscardReactionTimes.Checked && !this.isReactionTimeEntered(lane, true))
            {
                return false;
            }
            if (!this.isRiderSelected(lane, true))
            {
                return false;
            }
            if (!this.isRawTimeEntered(lane, true))
            {
                return false;
            }
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            _Worksheet wS = null;
            wS = this.getWorkSheet();
            string str = "";
            int num = 0;
            Range range = null;
            if (m_nRaceType == RaceTypes.Qualification)
            {
                range = wS.get_Range("C1", "C99");
                num = 1;
            }
            else
            {
                range = wS.get_Range("B1", "B99");
                num = 0;
            }
            Array values = (Array)range.Cells.Value2;
            string[] strArray = this.ConvertToStringArray(values);
            foreach (string str2 in strArray)
            {
                num++;
                if (str == "")
                {
                    if (m_nRaceType == RaceTypes.Qualification)
                    {
                        str = "C" + num.ToString();
                    }
                    else
                    {
                        string str3 = this.cbRound.SelectedItem.ToString();
                        if (str3 == "Final & consi")
                        {
                            string str4 = "";
                            if (lane == Lane.Left)
                            {
                                if (this.cbLeftRider.SelectedItem.ToString().Trim().Contains(" - (consi)"))
                                {
                                    str4 = "Cons";
                                }
                                else
                                {
                                    str4 = "Final";
                                }
                            }
                            else if (this.cbRigthRider.SelectedItem.ToString().Trim().Contains(" - (consi)"))
                            {
                                str4 = "Cons";
                            }
                            else
                            {
                                str4 = "Final";
                            }
                            if (str2.Contains(str4))
                            {
                                str = "B" + num.ToString();
                            }
                        }
                        else if (str2.IndexOf(str3) != -1)
                        {
                            str = "B" + num.ToString();
                        }
                    }
                    continue;
                }
                string str5 = "";
                if (m_nRaceType == RaceTypes.Qualification)
                {
                    str5 = "C" + num.ToString();
                }
                else
                {
                    str5 = "B" + num.ToString();
                }
                Range range2 = wS.get_Range(str5, str5);
                string str6 = "";
                if (lane == Lane.Left)
                {
                    str6 = this.cbLeftRider.SelectedItem.ToString().Trim();
                }
                else
                {
                    str6 = this.cbRigthRider.SelectedItem.ToString().Trim();
                }
                if (str6.EndsWith(" - (final)"))
                {
                    str6 = str6.Substring(0, str6.Length - 10);
                }
                if (str6.EndsWith(" - (consi)"))
                {
                    str6 = str6.Substring(0, str6.Length - 10);
                }
                if ((range2.Cells.Value2 != null) && ((range2.Cells.Value2.ToString().Trim() == str6) || (range2.Cells.Value2.ToString().Trim() == str6.Substring(str6.IndexOf(" ") + 1))))
                {
                    string str7 = "C" + num.ToString();
                    string str8 = "K" + num.ToString();
                    string cellCones = "D" + num.ToString();
                    string str10 = "L" + num.ToString();
                    string cellReaction = "E" + num.ToString();
                    string str12 = "M" + num.ToString();
                    string cellFalseStart = "F" + num.ToString();
                    string str14 = "N" + num.ToString();
                    if (m_nRaceType == RaceTypes.Qualification)
                    {
                        str7 = "D" + num.ToString();
                        cellCones = "E" + num.ToString();
                        cellReaction = "F" + num.ToString();
                        cellFalseStart = "G" + num.ToString();
                    }
                    string rawTime = "";
                    string reaction = "";
                    string cones = "";
                    if (lane == Lane.Left)
                    {
                        rawTime = this.tbLeftRaw.Text;
                        reaction = this.tbLeftReaction.Text;
                        cones = this.cbLeftCones.SelectedItem.ToString();
                    }
                    else
                    {
                        rawTime = this.tbRightRaw.Text;
                        reaction = this.tbRightReaction.Text;
                        cones = this.cbRightCones.SelectedItem.ToString();
                    }
                    Range range3 = wS.get_Range(str7, str7);
                    if ((range3.Cells.Value2 == null) || (range3.Cells.Value2.ToString() == ""))
                    {
                        this.SaveRunToExcel(wS, str7, rawTime, cellCones, cones, cellReaction, reaction, cellFalseStart, reaction.Replace("-", ""));
                        this.SaveAndReset(lane);
                        break;
                    }
                    range3 = wS.get_Range(str8, str8);
                    if ((range3.Cells.Value2 == null) || (range3.Cells.Value2.ToString() == ""))
                    {
                        this.SaveRunToExcel(wS, str8, rawTime, str10, cones, str12, reaction, str14, reaction.Replace("-", ""));
                        this.SaveAndReset(lane);
                        break;
                    }
                    RoundSelector selector = new RoundSelector(false);
                    if (selector.ShowDialog() != DialogResult.Cancel)
                    {
                        switch (selector.selectedRound)
                        {
                            case 1:
                                this.SaveRunToExcel(wS, str7, rawTime, cellCones, cones, cellReaction, reaction, cellFalseStart, reaction.Replace("-", ""));
                                break;

                            case 2:
                                this.SaveRunToExcel(wS, str8, rawTime, str10, cones, str12, reaction, str14, reaction.Replace("-", ""));
                                break;
                        }
                        this.SaveAndReset(lane);
                    }
                    break;
                }
            }
            Thread.CurrentThread.CurrentCulture = currentCulture;
            return true;
        }

        private bool saveDualQualification(Lane lane)
        {
            return this.saveDualLaneRace(lane);
        }

        private void SaveRunToExcel(_Worksheet WS, string cellRaw, string rawTime, string cellCones, string cones, string cellReaction, string reaction, string cellFalseStart, string falseStart)
        {
            WS.Cells.get_Range(cellRaw, cellRaw).Value2 = rawTime;
            WS.Cells.get_Range(cellCones, cellCones).Value2 = cones;
            if (!((bool)WS.Cells.get_Range(cellReaction, cellReaction).HasArray))
            {
                if (reaction.IndexOf("-") == -1)
                {
                    WS.Cells.get_Range(cellReaction, cellReaction).Value2 = reaction;
                    WS.Cells.get_Range(cellFalseStart, cellFalseStart).Value2 = "";
                }
                else
                {
                    WS.Cells.get_Range(cellReaction, cellReaction).Value2 = "";
                    WS.Cells.get_Range(cellFalseStart, cellFalseStart).Value2 = falseStart;
                }
            }
        }

        private void saveSingleLaneRace()
        {
            Lane left = Lane.Left;
            if ((this.isRiderSelected(left, true) && (this.cbDiscardReactionTimes.Checked || this.isReactionTimeEntered(left, true))) && this.isRawTimeEntered(left, true))
            {
                CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                string str = "";
                int num = 0;
                _Worksheet wS = null;
                wS = this.getWorkSheet();
                Range range = null;
                range = wS.get_Range("C1", "C99");
                num = 2;
                Array values = (Array)range.Cells.Value2;
                string[] strArray = this.ConvertToStringArray(values);
                foreach (string str2 in strArray)
                {
                    if (str2.ToLower().IndexOf("name") != -1)
                    {
                        str = "C" + num.ToString();
                    }
                    num++;
                    if (str != "")
                    {
                        string str3 = "";
                        str3 = "C" + num.ToString();
                        Range range2 = wS.get_Range(str3, str3);
                        string str4 = this.cbLeftRider.SelectedItem.ToString().Trim();
                        if ((range2.Cells.Value2 != null) && (range2.Cells.Value2.ToString().Trim() == str4))
                        {
                            string cellReaction = "D" + num.ToString();
                            string str6 = "E" + num.ToString();
                            string cellCones = "F" + num.ToString();
                            string str8 = "J" + num.ToString();
                            string str9 = "K" + num.ToString();
                            string str10 = "L" + num.ToString();
                            string str11 = "P" + num.ToString();
                            string str12 = "Q" + num.ToString();
                            string str13 = "R" + num.ToString();
                            string str14 = "V" + num.ToString();
                            string str15 = "W" + num.ToString();
                            string str16 = "X" + num.ToString();
                            string reaction = "";
                            string rawTime = "";
                            string cones = "";
                            reaction = this.tbLeftReaction.Text;
                            rawTime = this.tbLeftRaw.Text;
                            cones = this.cbLeftCones.SelectedItem.ToString();
                            Range range3 = wS.get_Range(str6, str6);
                            if ((range3.Cells.Value2 == null) || (range3.Cells.Value2.ToString() == ""))
                            {
                                this.SaveSmallRunToExcel(wS, str6, rawTime, cellCones, cones, cellReaction, reaction);
                                this.SaveAndReset(left);
                                break;
                            }
                            range3 = wS.get_Range(str9, str9);
                            if ((range3.Cells.Value2 == null) || (range3.Cells.Value2.ToString() == ""))
                            {
                                this.SaveSmallRunToExcel(wS, str9, rawTime, str10, cones, str8, reaction);
                                this.SaveAndReset(left);
                                break;
                            }
                            range3 = wS.get_Range(str12, str12);
                            if ((range3.Cells.Value2 == null) || (range3.Cells.Value2.ToString() == ""))
                            {
                                this.SaveSmallRunToExcel(wS, str12, rawTime, str13, cones, str11, reaction);
                                this.SaveAndReset(left);
                                break;
                            }
                            range3 = wS.get_Range(str15, str15);
                            if ((range3.Cells.Value2 == null) || (range3.Cells.Value2.ToString() == ""))
                            {
                                this.SaveSmallRunToExcel(wS, str15, rawTime, str16, cones, str14, reaction);
                                this.SaveAndReset(left);
                                break;
                            }
                            RoundSelector selector = new RoundSelector(true);
                            if (selector.ShowDialog() != DialogResult.Cancel)
                            {
                                switch (selector.selectedRound)
                                {
                                    case 1:
                                        this.SaveSmallRunToExcel(wS, str6, rawTime, cellCones, cones, cellReaction, reaction);
                                        break;

                                    case 2:
                                        this.SaveSmallRunToExcel(wS, str9, rawTime, str10, cones, str8, reaction);
                                        break;

                                    case 3:
                                        this.SaveSmallRunToExcel(wS, str12, rawTime, str13, cones, str11, reaction);
                                        break;

                                    case 4:
                                        this.SaveSmallRunToExcel(wS, str15, rawTime, str16, cones, str14, reaction);
                                        break;
                                }
                                this.SaveAndReset(left);
                            }
                            break;
                        }
                    }
                }
                this.SaveToExcel();
                Thread.CurrentThread.CurrentCulture = currentCulture;
            }
        }

        private void SaveSmallRunToExcel(_Worksheet WS, string cellRaw, string rawTime, string cellCones, string cones, string cellReaction, string reaction)
        {
            WS.Cells.get_Range(cellReaction, cellReaction).Value2 = reaction;
            WS.Cells.get_Range(cellRaw, cellRaw).Value2 = rawTime;
            WS.Cells.get_Range(cellCones, cellCones).Value2 = cones;
        }

        private void SaveToExcel()
        {
            try
            {
                this.theWorkbook.Save();
            }
            catch (Exception exception)
            {
                new MacMessageBox("It seems like the workbook have been opened in readonly mode. Please save this workbook manually and reopen this program with a non-read only version. " + exception.Message) { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
            }
        }

        private void setBGColor(Range range, int start, Color leftColor, Color rightColor)
        {
            int num = start;
            if ((num % 2) == 0)
            {
                range.Interior.Color = ColorTranslator.ToOle(leftColor);
            }
            else
            {
                range.Interior.Color = ColorTranslator.ToOle(rightColor);
            }
        }

        private void ShowFinalTimeOnDisplayLeft()
        {
            this.ShowFinalTimeOnDisplayLeft(null, null);
        }

        private void ShowFinalTimeOnDisplayLeft(object sender, EventArgs e)
        {
            if ((this.cbDisplays.Checked && (this.writeThreadLeft != null)) && this.writeThreadLeft.IsAlive)
            {
                this.StopDisplayThreadLeft();
                this.StartFinalTimeDisplayThreadLeft();
                this.StopDisplayThreadLeft();
            }
        }

        private void ShowFinalTimeOnDisplayRight()
        {
            this.ShowFinalTimeOnDisplayRight(null, null);
        }

        private void ShowFinalTimeOnDisplayRight(object sender, EventArgs e)
        {
            if ((this.cbDisplays.Checked && (this.writeThreadRight != null)) && this.writeThreadRight.IsAlive)
            {
                this.StopDisplayThreadRight();
                this.StartFinalTimeDisplayThreadRight();
                this.StopDisplayThreadRight();
            }
        }

        private void StartDisplayThreadLeft()
        {
            this.m_EventStopWriteToDisplayLeft.Reset();
            this.writeThreadLeft = new Thread(new ThreadStart(this.RunDisplayThreadLeft));
            this.writeThreadLeft.Name = "TrackMateDisplayWriterLeft";
            this.writeThreadLeft.Start();
        }

        private void StartDisplayThreadRight()
        {
            this.m_EventStopWriteToDisplayRight.Reset();
            this.writeThreadRight = new Thread(new ThreadStart(this.RunDisplayThreadRight));
            this.writeThreadRight.Name = "TrackMateDisplayWriterRight";
            this.writeThreadRight.Start();
        }

        private void StartFinalTimeDisplayThreadLeft()
        {
            this.m_EventStopWriteToDisplayLeft.Reset();
            this.writeThreadLeft = new Thread(new ThreadStart(this.RunFinalTimeDisplayThreadLeft));
            this.writeThreadLeft.Name = "TrackMateDisplayWriterLeft";
            this.writeThreadLeft.Start();
        }

        private void StartFinalTimeDisplayThreadRight()
        {
            this.m_EventStopWriteToDisplayRight.Reset();
            this.writeThreadRight = new Thread(new ThreadStart(this.RunFinalTimeDisplayThreadRight));
            this.writeThreadRight.Name = "TrackMateDisplayWriterRight";
            this.writeThreadRight.Start();
        }

        private void startThread()
        {
            if (this.readThread == null)
            {
                this.m_EventStopThread.Reset();
                this.readThread = new Thread(new ThreadStart(this.RunThread));
                this.readThread.Name = "TrackMateReaderThread";
                this.readThread.Start();
                this.bnConnect.Text = "Disconnect";
            }
        }

        private void StopDisplayThreadLeft()
        {
            if ((this.writeThreadLeft != null) && this.writeThreadLeft.IsAlive)
            {
                try
                {
                    this.m_EventStopWriteToDisplayLeft.Set();
                    while (this.writeThreadLeft.IsAlive)
                    {
                        Thread.Sleep(100);
                        System.Windows.Forms.Application.DoEvents();
                    }
                    this.writeThreadLeft.Join();
                    this.writeThreadLeft = null;
                }
                catch (Exception exception)
                {
                    string message = exception.Message;
                }
            }
        }

        private void StopDisplayThreadRight()
        {
            if ((this.writeThreadRight != null) && this.writeThreadRight.IsAlive)
            {
                try
                {
                    this.m_EventStopWriteToDisplayRight.Set();
                    while (this.writeThreadRight.IsAlive)
                    {
                        Thread.Sleep(100);
                        System.Windows.Forms.Application.DoEvents();
                    }
                    this.writeThreadRight.Join();
                    this.writeThreadRight = null;
                }
                catch (Exception exception)
                {
                    string message = exception.Message;
                }
            }
        }

        private void StopThread()
        {
            if ((this.readThread != null) && this.readThread.IsAlive)
            {
                this.m_EventStopThread.Set();
                while (this.readThread.IsAlive)
                {
                    Thread.Sleep(100);
                    System.Windows.Forms.Application.DoEvents();
                }
                this.readThread.Join();
                this.readThread = null;
            }
        }

        private void submitRunToWeb(string name, string start, string rawtime, string cones, string lane, bool countDown = false)
        {
            if (this.m_bolSendURL)
            {
                string str = "0";
                if (rawtime.ToLower().Contains("dq"))
                {
                    str = "1";
                }
                if (rawtime.Length == 0)
                {
                    rawtime = "0.00";
                }
                string text = "";
                try
                {
                    text = "http://slalomskateboarder.com/slalomranking/mvc.php?action=event.details&subaction=race_log&";
                    string str3 = "";
                    int num = 0;
                    if (name.EndsWith(" - (final)"))
                    {
                        str3 = name.Substring(0, name.Length - 10);
                        num++;
                    }
                    if (name.EndsWith(" - (consi)"))
                    {
                        str3 = name.Substring(0, name.Length - 10);
                    }
                    text = (text + "reporterid=" + this.tbReporterId.Text) + "&eventId=" + this.tbEventId.Text;
                    if (countDown)
                    {
                        text = text + "&callid=start_countdown";
                    }
                    else
                    {
                        text = ((((((text + "&racer=" + this.URLEncode(name)) + "&round=" + (this.getRoundId() + num)) + "&start=" + start.Replace(",", ".")) + "&time=" + rawtime.Replace(",", ".")) + "&cones=" + cones) + "&isDQ=" + str) + "&lane=" + lane;
                    }
                    if (MessageBox.Show(text, "URL", MessageBoxButtons.OKCancel) != DialogResult.Cancel)
                    {
                        for (int i = 0; i < 5; i++)
                        {
                            bool flag8 = false;
                            try
                            {
                                WebRequest request = WebRequest.Create(text);
                                request.Credentials = CredentialCache.DefaultCredentials;
                                ((HttpWebResponse)request.GetResponse()).Close();
                            }
                            catch (Exception exception)
                            {
                                string message = exception.Message;
                                flag8 = true;
                            }
                            if (!flag8)
                            {
                                return;
                            }
                        }
                    }
                }
                catch (Exception exception2)
                {
                    new MacMessageBox("Oops! The live reporting didn't work. Please check the internetconnection. " + exception2.Message) { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                }
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.cbPreviousData.Parent = this.tabRace;
            this.bnRefreshList.Parent = this.tabRace;
            if (this.cbLayOut.Checked)
            {
                Size size = new Size
                {
                    Height = 0x35,
                    Width = 0x284
                };
                this.groupLeft.Size = size;
                System.Drawing.Point location = new System.Drawing.Point
                {
                    X = 0,
                    Y = 0
                };
                this.groupLeft.Location = location;
                location.X = 6;
                location.Y = 15;
                this.cbLeftRider.Location = location;
                location.X = 370;
                location.Y = 15;
                this.tbLeftReaction.Location = location;
                location.X = 0x1c3;
                location.Y = 15;
                this.tbLeftRaw.Location = location;
                location.X = 0x221;
                location.Y = 14;
                this.cbLeftCones.Location = location;
                location.X = 0x290;
                location.Y = 8;
                this.bnSaveLeft.Location = location;
                location.X = 6;
                location.Y = 15;
                this.cbRigthRider.Location = location;
                location.X = 370;
                location.Y = 15;
                this.tbRightReaction.Location = location;
                location.X = 0x1c3;
                location.Y = 15;
                this.tbRightRaw.Location = location;
                location.X = 0x221;
                location.Y = 14;
                this.cbRightCones.Location = location;
                location.X = 0x290;
                location.Y = 0x3a;
                this.bnSaveRight.Location = location;
                location.X = 770;
                location.Y = 8;
                this.bnReset.Location = location;
                location.X = 770;
                location.Y = 0x54;
                this.bnRefreshList.Location = location;
                location.X = 0x16d;
                location.Y = -1;
                this.label7.Location = location;
                location.X = 0x1c3;
                location.Y = -1;
                this.label8.Location = location;
                location.X = 0x221;
                location.Y = -1;
                this.label9.Location = location;
                location.X = 0x16d;
                location.Y = -1;
                this.label14.Location = location;
                location.X = 0x1c3;
                location.Y = -1;
                this.label13.Location = location;
                location.X = 0x221;
                location.Y = -1;
                this.label6.Location = location;
                size = new Size
                {
                    Height = 0x35,
                    Width = 0x284
                };
                this.groupRight.Size = size;
                location = new System.Drawing.Point
                {
                    X = 0,
                    Y = 0x34
                };
                this.groupRight.Location = location;
                this.groupRight.Visible = true;
                this.groupLeft.Visible = true;
                this.label14.Visible = false;
                this.label13.Visible = false;
                this.label6.Visible = false;
                if (m_nRaceType == RaceTypes.SingleLane)
                {
                    this.groupRight.Visible = false;
                    this.label14.Visible = false;
                    this.label13.Visible = false;
                    this.label6.Visible = false;
                    this.cbRigthRider.Visible = false;
                    this.tbRightReaction.Visible = false;
                    this.tbRightRaw.Visible = false;
                    this.cbRightCones.Visible = false;
                    this.bnSaveRight.Visible = false;
                    this.groupLeft.Text = "Rider";
                    location.X = 770;
                    location.Y = 0x36;
                    this.bnRefreshList.Location = location;
                    location = this.cbPreviousData.Location;
                    location.X = 660;
                    location.Y = 0x36;
                    this.cbPreviousData.Location = location;
                    size = this.tbPrevData.Size;
                    size.Height = 0x9b;
                    this.tbPrevData.Size = size;
                    location.X = 5;
                    location.Y = 0x3a;
                    this.tbPrevData.Location = location;
                }
                else
                {
                    this.groupRight.Visible = true;
                    this.label14.Visible = true;
                    this.label13.Visible = true;
                    this.label6.Visible = true;
                    this.cbRigthRider.Visible = true;
                    this.tbRightReaction.Visible = true;
                    this.tbRightRaw.Visible = true;
                    this.cbRightCones.Visible = true;
                    this.groupLeft.Text = "Left";
                    location = this.cbPreviousData.Location;
                    location.X = 0x290;
                    location.Y = 0x57;
                    this.cbPreviousData.Location = location;
                    size = this.tbPrevData.Size;
                    size.Height = 110;
                    this.tbPrevData.Size = size;
                    location.X = 5;
                    location.Y = 0x69;
                    this.tbPrevData.Location = location;
                }
            }
            else if (m_nRaceType > RaceTypes.SingleLane)
            {
                this.bnRefreshList.Parent = this.groupRight;
                this.cbPreviousData.Parent = this.groupRight;
                Size size = new Size
                {
                    Height = 0x67,
                    Width = 0x177
                };
                this.groupLeft.Size = size;
                System.Drawing.Point location = new System.Drawing.Point
                {
                    X = 0,
                    Y = 0
                };
                this.groupLeft.Location = location;
                location = this.cbLeftRider.Location;
                location.X = 6;
                location.Y = 15;
                this.cbLeftRider.Location = location;
                location.X = 6;
                location.Y = 0x3f;
                this.tbLeftReaction.Location = location;
                location.X = 0x57;
                location.Y = 0x3f;
                this.tbLeftRaw.Location = location;
                location.X = 0xb5;
                location.Y = 0x3e;
                this.cbLeftCones.Location = location;
                location.X = 770;
                location.Y = 0x3d;
                this.bnSaveLeft.Location = location;
                location.X = 6;
                location.Y = 15;
                this.cbRigthRider.Location = location;
                location.X = 6;
                location.Y = 0x3f;
                this.tbRightReaction.Location = location;
                location.X = 0x57;
                location.Y = 0x3f;
                this.tbRightRaw.Location = location;
                location.X = 0xb5;
                location.Y = 0x3e;
                this.cbRightCones.Location = location;
                location.X = 770;
                location.Y = 0x3d;
                this.bnSaveRight.Location = location;
                location.X = 770;
                location.Y = 6;
                this.bnReset.Location = location;
                location.X = 0x291;
                location.Y = 0x34;
                this.bnRefreshList.Location = location;
                location.X = 5;
                location.Y = 0x31;
                this.label7.Location = location;
                location.X = 0x54;
                location.Y = 0x31;
                this.label8.Location = location;
                location.X = 0xb2;
                location.Y = 0x31;
                this.label9.Location = location;
                location.X = 5;
                location.Y = 0x31;
                this.label14.Location = location;
                location.X = 0x54;
                location.Y = 0x31;
                this.label13.Location = location;
                location.X = 0xb2;
                location.Y = 0x31;
                this.label6.Location = location;
                size = new Size
                {
                    Height = 0x67,
                    Width = 0x177
                };
                this.groupRight.Size = size;
                location = new System.Drawing.Point
                {
                    X = 390,
                    Y = 0
                };
                this.groupRight.Location = location;
                this.groupRight.Visible = true;
                this.groupLeft.Visible = true;
                this.label14.Visible = true;
                this.label13.Visible = true;
                this.label6.Visible = true;
                this.cbRigthRider.Visible = true;
                this.tbRightReaction.Visible = true;
                this.tbRightRaw.Visible = true;
                this.cbRightCones.Visible = true;
                location = this.cbPreviousData.Location;
                location.X = 0x106;
                location.Y = 80;
                this.cbPreviousData.Location = location;
                location = this.cbPreviousData.Location;
                location.X = 260;
                location.Y = 0x3d;
                this.bnRefreshList.Location = location;
                size = this.tbPrevData.Size;
                size.Height = 0x6c;
                this.tbPrevData.Size = size;
                location.X = 5;
                location.Y = 0x6b;
                this.tbPrevData.Location = location;
            }
            else
            {
                System.Drawing.Point location = this.cbLeftRider.Location;
                location.X = 6;
                location.Y = 15;
                this.cbLeftRider.Location = location;
                location.X = 370;
                location.Y = 15;
                this.tbLeftReaction.Location = location;
                location.X = 0x1c3;
                location.Y = 15;
                this.tbLeftRaw.Location = location;
                location.X = 0x221;
                location.Y = 14;
                this.cbLeftCones.Location = location;
                location.X = 0x290;
                location.Y = 8;
                this.bnSaveLeft.Location = location;
                this.groupRight.Visible = false;
                this.label14.Visible = false;
                this.label13.Visible = false;
                this.label6.Visible = false;
                this.cbRigthRider.Visible = false;
                this.tbRightReaction.Visible = false;
                this.tbRightRaw.Visible = false;
                this.cbRightCones.Visible = false;
                this.bnSaveRight.Visible = false;
                this.groupLeft.Text = "Rider";
                location.X = 770;
                location.Y = 0x36;
                this.bnRefreshList.Location = location;
                location = this.cbPreviousData.Location;
                location.X = 660;
                location.Y = 0x36;
                this.cbPreviousData.Location = location;
                Size size = this.tbPrevData.Size;
                size.Height = 0x9b;
                this.tbPrevData.Size = size;
                location.X = 5;
                location.Y = 0x3a;
                this.tbPrevData.Location = location;
            }
            if (this.cbColors.Checked)
            {
                if (this.cbRightColor.SelectedIndex != 0)
                {
                    this.cbRigthRider.BackColor = this.GetColor((string)this.cbRightColor.SelectedItem);
                    this.groupRight.BackColor = this.cbRigthRider.BackColor;
                }
                this.groupRight.Text = (string)this.cbRightColor.SelectedItem;
                if (this.cbLeftColor.SelectedIndex != 0)
                {
                    this.cbLeftRider.BackColor = this.GetColor((string)this.cbLeftColor.SelectedItem);
                    this.groupLeft.BackColor = this.cbLeftRider.BackColor;
                }
                this.groupLeft.Text = (string)this.cbLeftColor.SelectedItem;
            }
            if (((this.tabControl1.SelectedIndex != 0) && this.cbDisplays.Checked) && (((this.cbComPort.Text == this.cbDisplayPortLeft.Text) || (this.cbComPort.Text == this.cbDisplayPortRight.Text)) || (this.cbDisplayPortLeft.Text == this.cbDisplayPortRight.Text)))
            {
                new MacMessageBox("You need to select different comports for each connection if you're using external displays.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                this.tabControl1.SelectedIndex = 0;
            }
            if ((this.tabControl1.SelectedIndex != 0) && ((this.tbFileName.Text.Length == 0) || (this.theWorkbook == null)))
            {
                new MacMessageBox("Please select an Excel file.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                this.tabControl1.SelectedIndex = 0;
            }
            if ((this.tabControl1.SelectedIndex != 0) && (m_nRaceType == RaceTypes.NotSet))
            {
                new MacMessageBox("Please select worksheet.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                this.tabControl1.SelectedIndex = 0;
            }
            if (((this.tabControl1.SelectedIndex != 0) && (m_nRaceType == RaceTypes.Elimination)) && (this.cbRound.SelectedIndex == -1))
            {
                new MacMessageBox("Please select which round of the race we are running.") { StartPosition = FormStartPosition.CenterParent }.ShowDialog();
                this.tabControl1.SelectedIndex = 0;
            }
            if (this.cbDiscardReactionTimes.Checked)
            {
                this.tbRightReaction.BackColor = Color.Gray;
                this.tbLeftReaction.BackColor = Color.Gray;
            }
            else
            {
                this.tbRightReaction.BackColor = Color.White;
                this.tbLeftReaction.BackColor = Color.White;
            }
            if (this.tabControl1.SelectedIndex == 0)
            {
                base.Height = 0xa1;
            }
            else if (m_nRaceType == RaceTypes.SingleLane)
            {
                base.Height = 0x83;
            }
            else
            {
                base.Height = 0xa1;
            }
        }

        private string URLEncode(string text)
        {
            char[] chArray = text.ToCharArray();
            string str = "";
            foreach (char ch in chArray)
            {
                if (ch > '\x0080')
                {
                    str = str + "%" + ((int)ch).ToString("X");
                }
                else
                {
                    str = str + ch.ToString();
                }
            }
            return str;
        }

        private void WriteToLogfile(string message)
        {
            try
            {
                if (this.cbLog2File.Checked)
                {
                    if (!System.IO.File.Exists(this.m_strLogFile))
                    {
                        System.IO.File.Create(this.m_strLogFile);
                    }
                    StreamWriter writer = new StreamWriter(this.m_strLogFile, true);
                    writer.WriteLine(message + " --- " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    writer.Flush();
                    writer.Close();
                }
            }
            catch (Exception)
            {
            }
        }

        private enum Lane
        {
            Left = 1,
            Right = 2
        }

        private enum RaceTypes
        {
            NotSet = -1,
            SingleLane = 0,
            Qualification = 1,
            Elimination = 2
        }

        public delegate void TrackMateCallback(string message, int lane, int type);
    }
}
