using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Mac.Excel9.Interop;
using System.Threading;
using System.IO.Ports;
using System.Net;
using System.IO;

namespace ExcelMate
{
    public partial class mExcelMate : Form
    {

        //
        // TODO:
        //
        //

        #region Public Delegates
        public delegate void TrackMateCallback(string message, int lane, int type);
        #endregion

        #region variables
        Thread readThread;
        Thread writeThreadLeft;
        Thread writeThreadRight;
        ManualResetEvent m_EventStopThread;
        ManualResetEvent m_EventResetTimer;
        ManualResetEvent m_EventStopWriteToDisplayLeft;
        ManualResetEvent m_EventStopWriteToDisplayRight;
        public TrackMateCallback mTrackMateCallback;

        private Mac.Excel9.Interop.Application objExcel = null;
        private Workbook theWorkbook = null;
        private static RaceTypes m_nRaceType = 0; // 0==single lane, 1==dual qual, 2 == dual elimination 
        private String m_strReportSiteSelection = "";
        private Boolean m_bolTrackmateVersion = false;
        private Boolean m_bolDisplayZeroReaction = false;
        private String m_strWorkBookName = "";
        private string m_strLogFile = "";

        #endregion variables

        #region enums

        private enum RaceTypes{ NotSet = -1, SingleLane = 0, Qualification = 1, Elimination = 2};
        private enum Lane { Left = 1, Right = 2 };

        #endregion enums

        public mExcelMate()
        {
            InitializeComponent();
            try
            {
                objExcel = new Mac.Excel9.Interop.Application();
            }
            catch (Exception ex)
            {
                MacMessageBox MMB = new MacMessageBox("The startup of this application failed, probably because you do not have Excel installed on the computer.", MessageBoxButtons.OK);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
                return;
            }
            mTrackMateCallback = new TrackMateCallback(this.AddTrackMateMessage);
            m_EventStopThread = new ManualResetEvent(false);
            m_EventResetTimer = new ManualResetEvent(false);

            m_EventStopWriteToDisplayLeft = new ManualResetEvent(false);
            m_EventStopWriteToDisplayRight = new ManualResetEvent(false);

            foreach (string s in SerialPort.GetPortNames())
            {
                cbComPort.Items.Add(s);
                cbDisplayPortLeft.Items.Add(s);
                cbDisplayPortRight.Items.Add(s);
            }

            cbComPort.SelectedIndex = -1;
            cbDisplayPortLeft.SelectedIndex = -1;
            cbDisplayPortRight.SelectedIndex = -1;

            cbDisplays.Checked = false;

            tbLeftRaw.Text = "";
            tbLeftReaction.Text = "";
            cbLeftCones.SelectedIndex = 0;
            cbLeftRider.SelectedIndex = -1;
            tbRightRaw.Text = "";
            tbRightReaction.Text = "";
            cbRightCones.SelectedIndex = 0;
            cbRigthRider.SelectedIndex = -1;
            m_strLogFile = "RaceLog_" + DateTime.Today.ToString("yyyyMMdd") + ".log";
            cbRightColor.Enabled = false;
            cbLeftColor.Enabled = false;
            cbRightColor.SelectedIndex = 0;
            cbLeftColor.SelectedIndex = 0;
        }

        #region ExcelStuff

        private void bnOpenWorkBook_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            m_nRaceType = RaceTypes.NotSet;
            this.openFileDialog1.FileName = "*.xls";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbFileName.Text = openFileDialog1.FileName;

                try
                {
                    System.Globalization.CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
                    Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                    // close old stuff
                    if (theWorkbook != null)
                    {
                        theWorkbook.Save();
                        theWorkbook.Close(true, m_strWorkBookName, false);
                        objExcel.Visible = false;
                        objExcel = null;
                        Thread.Sleep(1000);
                        objExcel = new Mac.Excel9.Interop.Application();
                    }

                    theWorkbook = objExcel.Workbooks.Open(tbFileName.Text, 0, false, 5, "", "", true, Mac.Excel9.Interop.XlPlatform.xlWindows, "\t", false, false, 0, true);

                    Thread.CurrentThread.CurrentCulture = oldCI;
                }
                catch (Exception ex)
                {
                    MacMessageBox MMB = new MacMessageBox("I guess you were not able to open the workbook. Please make sure it is available and not readonly. "+ ex.Message);
                    MMB.StartPosition = FormStartPosition.CenterParent;
                    MMB.ShowDialog();
                    tbFileName.Text = "";
                    return;
                }

                // get the collection of sheets in the workbook
                Mac.Excel9.Interop.Sheets sheets = theWorkbook.Worksheets;
                cbWorkSheet.Items.Clear();

                for (int i = 1; i <= sheets.Count; i++)
                {
                    Mac.Excel9.Interop._Worksheet WS = (Mac.Excel9.Interop._Worksheet)sheets.get_Item(i);
                    cbWorkSheet.Items.Add(WS.Name);
                }
                this.Text = tbFileName.Text;
                m_strWorkBookName = tbFileName.Text;
            }
            this.Cursor = Cursors.Default;
        }

        private void cbWorkSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            Mac.Excel9.Interop._Worksheet WS = null;

            try{
                bnRefreshList.Visible = true;
                System.Globalization.CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                WS = (Mac.Excel9.Interop._Worksheet)theWorkbook.Worksheets.get_Item(cbWorkSheet.SelectedIndex + 1);
                WS.Activate();
                objExcel.Visible = true;
                this.TopMost = true;
                this.TopMost = false;
                cbRound.Items.Clear();

                // this is qualification or single lane
                if (WS.Name.ToLower().IndexOf("qual") != -1){
                    MacMessageBox MMB = new MacMessageBox("This race will be treated as qualification in head 2 head format (since the workbook contains 'qual').");
                    MMB.StartPosition = FormStartPosition.CenterParent;
                    MMB.ShowDialog();
                    m_nRaceType = RaceTypes.Qualification;
                }else if(WS.Name.ToLower().IndexOf("elim") != -1){
                    MacMessageBox MMB = new MacMessageBox("This race will be treated as eliminatin in head 2 head format (since the workbook contains 'elim').");
                    MMB.StartPosition = FormStartPosition.CenterParent;
                    MMB.ShowDialog();
                    m_nRaceType = RaceTypes.Elimination;
                }else{
                    MacMessageBox MMB = new MacMessageBox("This race will be treated as single lane (since the workbook does not contain 'qual' or 'elim').");
                    MMB.StartPosition = FormStartPosition.CenterParent;
                    MMB.ShowDialog();
                    m_nRaceType = RaceTypes.SingleLane;
                }

                if (m_nRaceType == RaceTypes.SingleLane || m_nRaceType == RaceTypes.Qualification){
                    FillRiderDropDownsSingle(WS);
                }
                else
                {
                    cbRound.Enabled = true;
                    bnRefreshList.Visible = false;

                    Mac.Excel9.Interop.Range range = WS.get_Range("B1", "B99");
                    System.Array myvalues = (System.Array)range.Cells.Value2;
                    string[] strArray = ConvertToStringArray(myvalues);

                    foreach (String cellValue in strArray)
                    {
                        if (cellValue.ToLower().IndexOf("round") != -1)
                        {
                            if (cellValue.ToLower().IndexOf("final") != -1)
                            {
                                cbRound.Items.Add("Final & consi");
                            }
                            else if (cellValue.ToLower().IndexOf("cons") != -1)
                            {
                            }
                            else
                            {
                                cbRound.Items.Add(cellValue);
                            }
                        }
                    }

                    if (cbRound.Items.Count == 0)
                    {
                        MacMessageBox MMB = new MacMessageBox("No racers was found in the Excel sheet. If this is single lane or qualification remember that \r\nthe name of the workbook must contain 'qual' in order for the program to find the racers.");
                        MMB.StartPosition = FormStartPosition.CenterParent;
                        MMB.ShowDialog();
                    }
                }
                Thread.CurrentThread.CurrentCulture = oldCI;
            }
            catch (Exception ex)
            {
                MacMessageBox MMB = new MacMessageBox("Oops! We were not able to read the excelsheet that you previously selected. Did you kill it? Please re-open the spread sheet. "+ ex.Message);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
            }
        }

        private void FillRiderDropDownsSingle(Mac.Excel9.Interop._Worksheet WS)
        {
            cbRound.Enabled = false;
            cbLeftRider.Items.Clear();
            cbRigthRider.Items.Clear();
            Range range = WS.get_Range("C1", "C99");
            Array myvalues = (Array)range.Cells.Value2;
            string[] strArray = ConvertToStringArray(myvalues);
            string strStartRound = "";
            int nStartValue = 2;
            foreach (String cellValue in strArray)
            {
                if (cellValue.ToLower().IndexOf("name") != -1)
                {
                    nStartValue++;
                    strStartRound = "C" + nStartValue.ToString();
                }
                else
                {
                    nStartValue++;
                }

                if (strStartRound != "")
                {
                    string strCellRiderName = "C" + nStartValue.ToString();
                    Range range0 = WS.get_Range(strCellRiderName, strCellRiderName);
                    if (range0.Value2 == null)
                    {
                        break;
                    }
                    string strRiderName = range0.Cells.Value2.ToString();
                    cbLeftRider.Items.Add(strRiderName);
                    cbRigthRider.Items.Add(strRiderName);
                }
            }
            tabControl1.SelectedIndex = 1;
        }

        string[] ConvertToStringArray(System.Array values)
        {
            // create a new string array
            string[] theArray = new string[values.Length];
            // loop through the 2-D System.Array and populate the 1-D String Array
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(i, 1) == null)
                    theArray[i - 1] = "";
                else
                    theArray[i - 1] = (string)values.GetValue(i, 1).ToString();
            }
            return theArray;
        }

        private void bnSaveLeft_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            switch (m_nRaceType)
            {
                case RaceTypes.SingleLane:
                    {
                        saveSingleLaneRace();
                        break;
                    }
                case RaceTypes.Qualification:
                    {
                        bool save = false;
                        save = saveDualQualification(Lane.Left);
                        save = saveDualQualification(Lane.Right);
                        if (save)
                        {
                            SaveToExcel();
                        }
                        break;
                    }
                case RaceTypes.Elimination:
                    {
                        bool save = false;
                        save = saveDualLaneRace(Lane.Left);
                        save = saveDualLaneRace(Lane.Right);
                        if (save)
                        {
                            SaveToExcel();
                        }
                        break;
                    }
            }
            this.Cursor = Cursors.Default;
        }

        private Boolean isRiderSelected(Lane lane, Boolean showMsg)
        {
            if (lane == Lane.Left)
            {
                if (cbLeftRider.SelectedIndex != -1){
                    return true;
                }
            }
            else
            {
                if (cbRigthRider.SelectedIndex != -1){
                    return true;
                }
            }

            if (showMsg)
            {
                MacMessageBox MMB = new MacMessageBox("Please select a rider to save the time to.");
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
            }
            return false;
        }

        private Boolean isRawTimeEntered(Lane lane, Boolean showMsg)
        {
            if (lane == Lane.Left)
            {
                if (tbLeftRaw.Text.Length != 0)
                {
                    return true;
                }
            }
            else
            {
                if (tbRightRaw.Text.Length != 0)
                {
                    return true;
                }
            }

            if (showMsg)
            {
                MacMessageBox MMB = new MacMessageBox("There is no raw time to save.");
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
            }
            return false;
        }

        private Boolean isReactionTimeEntered(Lane lane, Boolean showMsg)
        {
            if (lane == Lane.Left)
            {
                if (tbLeftReaction.Text.Length != 0)
                {
                    return true;
                }
            }
            else
            {
                if (tbRightReaction.Text.Length != 0)
                {
                    return true;
                }
            }

            if (showMsg)
            {
                MacMessageBox MMB = new MacMessageBox("There is no reaction time to save.");
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
            }
            return false;
        }

        private bool saveDualQualification(Lane lane){
            return saveDualLaneRace(lane);
        }

        private _Worksheet getWorkSheet()
        {
            _Worksheet WS = null;
            try
            {
                WS = (_Worksheet)theWorkbook.Worksheets.get_Item(cbWorkSheet.SelectedIndex + 1);
            }
            catch (Exception ex)
            {
                MacMessageBox MMB = new MacMessageBox("Oops! We were not able to read the excelsheet that you previously selected. Did you kill it? Please re-open the spread sheet. " + ex.Message);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
            }
            return WS;
        }

        private bool saveDualLaneRace(Lane lane){
            // We would like to find the selected rider and make sure that there is no
            // time saved. If there are times save already we should either move on to round
            // 2 and make sure that no data have been saved in that field and ask if
            // is correct that this is round 2. If there is data saved in round 2 as well then
            // we ask if this is a rerun and which part we should overwrite.
            if (!isRiderSelected(lane, false) && !isRawTimeEntered(lane, false) && !isReactionTimeEntered(lane, false))
            {
                return false;
            }

            if (!cbDiscardReactionTimes.Checked)
            {
                if (!isReactionTimeEntered(lane, true))
                {
                    return false;
                }
            }
            
            if (!isRiderSelected(lane, true))
            {
                return false;
            }
            if (!isRawTimeEntered(lane, true))
            {
                return false;
            }

            System.Globalization.CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            Mac.Excel9.Interop._Worksheet WS = null;
            WS = getWorkSheet();
            string strStartRound = "";
            int nStartValue = 0;

            // retrieve the riders that are in this round
            Mac.Excel9.Interop.Range range = null;

            // 1 == dual qual 
            if (m_nRaceType == RaceTypes.Qualification)
            {
                range = WS.get_Range("C1", "C99");
                nStartValue = 1;
            }
            else
            {
                range = WS.get_Range("B1", "B99");
                nStartValue = 0;
            }

            System.Array myvalues = (System.Array)range.Cells.Value2;
            string[] strArray = ConvertToStringArray(myvalues);

            // loop until you find where the selected round starts and 
            // then gather the riders within this round.
            foreach (String cellValue in strArray)
            {
                nStartValue++;
                if (strStartRound == "")
                {
                    // 1 == dual qual 
                    if (m_nRaceType == RaceTypes.Qualification)
                    {
                        strStartRound = "C" + nStartValue.ToString();
                    }
                    else
                    {
                        String strSelRound = cbRound.SelectedItem.ToString();
                        if (strSelRound == "Final & consi")
                        {
                            string strWhatFinal = "";
                            if (lane == Lane.Left)
                            {
                                if (cbLeftRider.SelectedItem.ToString().Trim().Contains(" - (consi)"))
                                {
                                    strWhatFinal = "Cons";
                                }
                                else
                                {
                                    strWhatFinal = "Final";
                                }
                            }
                            else
                            {
                                if (cbRigthRider.SelectedItem.ToString().Trim().Contains(" - (consi)"))
                                {
                                    strWhatFinal = "Cons";
                                }
                                else
                                {
                                    strWhatFinal = "Final";
                                }
                            }

                            if (cellValue.Contains(strWhatFinal))
                            {
                                strStartRound = "B" + nStartValue.ToString();
                            }
                        }
                        else
                        {
                            if (cellValue.IndexOf(strSelRound) != -1)
                            {
                                strStartRound = "B" + nStartValue.ToString();
                            }
                        }
                    }
                }else{
                    // find the selected rider and check if there are times saved already.
                    string strCellRiderName = "";

                    // 1 == dual qual 
                    if (m_nRaceType == RaceTypes.Qualification)
                    {
                        strCellRiderName = "C" + nStartValue.ToString();
                    }
                    else
                    {
                        strCellRiderName = "B" + nStartValue.ToString();
                    }
                    Mac.Excel9.Interop.Range range0 = WS.get_Range(strCellRiderName, strCellRiderName);
                    string selected = "";
                    if (lane == Lane.Left)
                    {
                        selected = cbLeftRider.SelectedItem.ToString().Trim(); 
                    }
                    else
                    {
                        selected = cbRigthRider.SelectedItem.ToString().Trim();
                    }

                    if (selected.EndsWith(" - (final)"))
                    {
                        selected = selected.Substring(0, selected.Length - 10);
                    }
                    if (selected.EndsWith(" - (consi)"))
                    {
                        selected = selected.Substring(0, selected.Length - 10);
                    }

                    if (range0.Cells.Value2 != null && range0.Cells.Value2.ToString().Trim() == selected)
                    {
                        string strCellFirstTime = "C" + nStartValue.ToString();
                        string strCellSecondTime = "K" + nStartValue.ToString();
                        string strCellFirstCones = "D" + nStartValue.ToString();
                        string strCellSecondCones = "L" + nStartValue.ToString();
                        string strCellFirstReaction = "E" + nStartValue.ToString();
                        string strCellSecondReaction = "M" + nStartValue.ToString();
                        string strCellFirstFalse = "F" + nStartValue.ToString();
                        string strCellSecondFalse = "N" + nStartValue.ToString();

                        // 1 == dual qual 
                        if (m_nRaceType == RaceTypes.Qualification)
                        {
                            strCellFirstTime = "D" + nStartValue.ToString();
                            strCellFirstCones = "E" + nStartValue.ToString(); ;
                            strCellFirstReaction = "F" + nStartValue.ToString();
                            strCellFirstFalse = "G" + nStartValue.ToString();
                        }

                        string strRaw = "";
                        string strReaction = "";
                        string strCones = "";

                        if (lane == Lane.Left)
                        {
                            strRaw = tbLeftRaw.Text;
                            strReaction = tbLeftReaction.Text;
                            strCones = cbLeftCones.SelectedItem.ToString();
                        }
                        else
                        {
                            strRaw = tbRightRaw.Text;
                            strReaction = tbRightReaction.Text;
                            strCones = cbRightCones.SelectedItem.ToString();
                        }

                        // SaveData as First run
                        Mac.Excel9.Interop.Range range1 = WS.get_Range(strCellFirstTime, strCellFirstTime);
                        if (range1.Cells.Value2 == null || range1.Cells.Value2.ToString() == "")
                        {
                            SaveRunToExcel(WS, strCellFirstTime, strRaw, strCellFirstCones, strCones, strCellFirstReaction, strReaction, strCellFirstFalse, strReaction.Replace("-", ""));
                            SaveAndReset(lane);
                            break;
                        }
                        else
                        {
                            // SaveData as second run
                            range1 = WS.get_Range(strCellSecondTime, strCellSecondTime);
                            if (range1.Cells.Value2 == null || range1.Cells.Value2.ToString() == "")
                            {
                                SaveRunToExcel(WS, strCellSecondTime, strRaw, strCellSecondCones, strCones, strCellSecondReaction, strReaction, strCellSecondFalse, strReaction.Replace("-", ""));
                                SaveAndReset(lane);
                                break;
                            }
                            else
                            {
                                RoundSelector RS = new RoundSelector(false);
                                if (RS.ShowDialog() != DialogResult.Cancel)
                                {
                                    switch (RS.selectedRound)
                                    {
                                        case 1:
                                            SaveRunToExcel(WS, strCellFirstTime, strRaw, strCellFirstCones, strCones, strCellFirstReaction, strReaction, strCellFirstFalse, strReaction.Replace("-", ""));
                                            break;
                                        case 2:
                                            SaveRunToExcel(WS, strCellSecondTime, strRaw, strCellSecondCones, strCones, strCellSecondReaction, strReaction, strCellSecondFalse, strReaction.Replace("-", ""));
                                            break;
                                    }
                                    SaveAndReset(lane);
                                }
                                break;
                            }
                        }
                    }
                }
            }
            Thread.CurrentThread.CurrentCulture = oldCI;
            return true;
        }

        private void saveSingleLaneRace(){
            // We would like to find the selected rider and make sure that there is no
            // time saved. If there are times save already we should either move on to round
            // 2 and make sure that no data have been saved in that field and ask if
            // is correct that this is round 2. If there is data saved in round 2 as well then
            // we ask if this is a rerun and which part we should overwrite.
            Lane lane = Lane.Left;
            if (!isRiderSelected(lane, true))
            {
                return;
            }
            if (!cbDiscardReactionTimes.Checked)
            {
                if (!isReactionTimeEntered(lane, true))
                {
                    return;
                }
            }
            if (!isRawTimeEntered(lane, true))
            {
                return;
            }

            System.Globalization.CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            string strStartRound = "";
            int nStartValue = 0;

            Mac.Excel9.Interop._Worksheet WS = null;
            WS = getWorkSheet();

            // retrieve the riders that are in this round
            Mac.Excel9.Interop.Range range = null;

            range = WS.get_Range("C1", "C99");
            nStartValue = 2;

            System.Array myvalues = (System.Array)range.Cells.Value2;
            string[] strArray = ConvertToStringArray(myvalues);

            // loop until you find where the selected round starts and 
            // then gather the riders within this round.
            foreach (String cellValue in strArray)
            {
                if (cellValue.ToLower().IndexOf("name") != -1)
                {
                    strStartRound = "C" + nStartValue.ToString();
                }

                nStartValue++;

                if (strStartRound != "")
                {
                    // find the selected rider and check if there are times saved already.
                    string strCellRiderName = "";

                    strCellRiderName = "C" + nStartValue.ToString();

                    Mac.Excel9.Interop.Range range0 = WS.get_Range(strCellRiderName, strCellRiderName);
                    string selected = cbLeftRider.SelectedItem.ToString().Trim();

                    if (range0.Cells.Value2 != null && range0.Cells.Value2.ToString().Trim() == selected)
                    {
                        string strCellFirstReaction = "D" + nStartValue.ToString();
                        string strCellFirstTime = "E" + nStartValue.ToString();
                        string strCellFirstCones = "F" + nStartValue.ToString(); ;
                        string strCellSecondReaction = "J" + nStartValue.ToString();
                        string strCellSecondTime = "K" + nStartValue.ToString();
                        string strCellSecondCones = "L" + nStartValue.ToString();
                        string strCellThirdReaction = "P" + nStartValue.ToString();
                        string strCellThirdTime = "Q" + nStartValue.ToString();
                        string strCellThirdCones = "R" + nStartValue.ToString();
                        string strCellFourthReaction = "V" + nStartValue.ToString();
                        string strCellFourthTime = "W" + nStartValue.ToString();
                        string strCellFourthCones = "X" + nStartValue.ToString();

                        string strReaction = "";
                        string strRaw = "";
                        string strCones = "";

                        strReaction = tbLeftReaction.Text;
                        strRaw = tbLeftRaw.Text;
                        strCones = cbLeftCones.SelectedItem.ToString();

                        // SaveData as First run
                        Mac.Excel9.Interop.Range range1 = WS.get_Range(strCellFirstTime, strCellFirstTime);
                        if (range1.Cells.Value2 == null || range1.Cells.Value2.ToString() == "")
                        {
                            SaveSmallRunToExcel(WS, strCellFirstTime, strRaw, strCellFirstCones, strCones, strCellFirstReaction, strReaction);
                            SaveAndReset(lane);
                            break;
                        }
                        else
                        {
                            range1 = WS.get_Range(strCellSecondTime, strCellSecondTime);
                            if (range1.Cells.Value2 == null || range1.Cells.Value2.ToString() == "")
                            {
                                SaveSmallRunToExcel(WS, strCellSecondTime, strRaw, strCellSecondCones, strCones, strCellSecondReaction, strReaction);
                                SaveAndReset(lane);
                                break;
                            }
                            else
                            {
                                range1 = WS.get_Range(strCellThirdTime, strCellThirdTime);
                                if (range1.Cells.Value2 == null || range1.Cells.Value2.ToString() == "")
                                {
                                    SaveSmallRunToExcel(WS, strCellThirdTime, strRaw, strCellThirdCones, strCones, strCellThirdReaction, strReaction);
                                    SaveAndReset(lane);
                                    break;
                                }
                                else
                                {
                                    range1 = WS.get_Range(strCellFourthTime, strCellFourthTime);
                                    if (range1.Cells.Value2 == null || range1.Cells.Value2.ToString() == "")
                                    {
                                        SaveSmallRunToExcel(WS, strCellFourthTime, strRaw, strCellFourthCones, strCones, strCellFourthReaction, strReaction);
                                        SaveAndReset(lane);
                                        break;
                                    }
                                    else
                                    {
                                        RoundSelector RS = new RoundSelector(true);

                                        if (RS.ShowDialog() != DialogResult.Cancel)
                                        {
                                            switch (RS.selectedRound)
                                            {
                                                case 1:
                                                    SaveSmallRunToExcel(WS, strCellFirstTime, strRaw, strCellFirstCones, strCones, strCellFirstReaction, strReaction);
                                                    break;
                                                case 2:
                                                    SaveSmallRunToExcel(WS, strCellSecondTime, strRaw, strCellSecondCones, strCones, strCellSecondReaction, strReaction);
                                                    break;
                                                case 3:
                                                    SaveSmallRunToExcel(WS, strCellThirdTime, strRaw, strCellThirdCones, strCones, strCellThirdReaction, strReaction);
                                                   break;
                                                case 4:
                                                   SaveSmallRunToExcel(WS, strCellFourthTime, strRaw, strCellFourthCones, strCones, strCellFourthReaction, strReaction);
                                                    break;
                                            }
                                            SaveAndReset(lane);
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            SaveToExcel();
            Thread.CurrentThread.CurrentCulture = oldCI;
        }

        private void SaveRunToExcel(Mac.Excel9.Interop._Worksheet WS, String cellRaw, String rawTime, String cellCones, String cones, String cellReaction, String reaction, String cellFalseStart, String falseStart)
        {
            ((Mac.Excel9.Interop.Range)WS.Cells.get_Range(cellRaw, cellRaw)).Value2 = rawTime;
            ((Mac.Excel9.Interop.Range)WS.Cells.get_Range(cellCones, cellCones)).Value2 = cones;
            if (!(Boolean)((Mac.Excel9.Interop.Range)WS.Cells.get_Range(cellReaction, cellReaction)).HasArray)
            {
                if (reaction.IndexOf("-") == -1)
                {
                    ((Mac.Excel9.Interop.Range)WS.Cells.get_Range(cellReaction, cellReaction)).Value2 = reaction;
                    ((Mac.Excel9.Interop.Range)WS.Cells.get_Range(cellFalseStart, cellFalseStart)).Value2 = "";
                }
                else
                {
                    ((Mac.Excel9.Interop.Range)WS.Cells.get_Range(cellReaction, cellReaction)).Value2 = "";
                    ((Mac.Excel9.Interop.Range)WS.Cells.get_Range(cellFalseStart, cellFalseStart)).Value2 = falseStart;
                }
            }
        }

        private void SaveSmallRunToExcel(Mac.Excel9.Interop._Worksheet WS, String cellRaw, String rawTime, String cellCones, String cones, String cellReaction, String reaction)
        {
            ((Mac.Excel9.Interop.Range)WS.Cells.get_Range(cellReaction, cellReaction)).Value2 = reaction;
            ((Mac.Excel9.Interop.Range)WS.Cells.get_Range(cellRaw, cellRaw)).Value2 = rawTime;
            ((Mac.Excel9.Interop.Range)WS.Cells.get_Range(cellCones, cellCones)).Value2 = cones;
        }

        private void SaveToExcel()
        {
            try
            {
                theWorkbook.Save();
            }
            catch (Exception ex)
            {
                MacMessageBox MMB = new MacMessageBox("It seems like the workbook have been opened in readonly mode. Please save this workbook manually and reopen this program with a non-read only version. " + ex.Message);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
            }
        }

        private void SaveAndReset(Lane lane)
        {
            String strRider;
            String strReaction;
            String strRaw;
            String strCones;
            String strLane;
            if (lane == Lane.Left)
            {
                strRider = cbLeftRider.SelectedItem.ToString().Trim();
                strReaction = tbLeftReaction.Text;
                strRaw = tbLeftRaw.Text.Replace(".", ",");
                strCones = cbLeftCones.SelectedItem.ToString();
                strLane = "R";
            }
            else
            {
                strRider = cbRigthRider.SelectedItem.ToString().Trim();
                strReaction = tbRightReaction.Text;
                strRaw = tbRightRaw.Text.Replace(".", ",");
                strCones = cbRightCones.SelectedItem.ToString();
                strLane = "W";
            }

            if (m_nRaceType == RaceTypes.SingleLane)
            {
                strLane = "W";
            }

            if (cbLiveReport.Checked)
            {
                if (tbLiveId.Text.Length == 0)
                {
                    MacMessageBox MMB = new MacMessageBox("You need to fill in the Live Reporter Id in order to submit times to the web. Get your own reporter Id, send an email to marcus@ettsexett.com.");
                    MMB.StartPosition = FormStartPosition.CenterParent;
                    MMB.ShowDialog();
                    return;
                }
                if (tbLiveEventId.Text.Length == 0)
                {
                    MacMessageBox MMB = new MacMessageBox("You need to fill in the EventId in order to submit times to the web. To get the event Id, send an email to marcus@ettsexett.com.");
                    MMB.StartPosition = FormStartPosition.CenterParent;
                    MMB.ShowDialog();
                    return;
                }

                submitRunToWeb(strRider, strReaction, strRaw, strCones, strLane);
            }
            tbPrevData.Text = strRider + " - " + strReaction + ", " + strRaw + " + " + strCones + Environment.NewLine + tbPrevData.Text;
            if (lane == Lane.Left)
            {
                tbLeftReaction.Text = "";
                tbLeftReaction.Refresh();
                tbLeftRaw.Text = "";
                tbLeftRaw.Refresh();
                cbLeftCones.SelectedIndex = 0;
                if (m_nRaceType == RaceTypes.SingleLane) // single lane
                {
                    if (cbLeftRider.Items.Count > cbLeftRider.SelectedIndex + 1)
                    {
                        cbLeftRider.SelectedIndex = cbLeftRider.SelectedIndex + 1;
                    }
                    else
                    {
                        cbLeftRider.SelectedIndex = -1;
                    }
                }
                else
                {
                    if (cbLeftRider.Items.Count > cbLeftRider.SelectedIndex + 2)
                    {
                        cbLeftRider.SelectedIndex = cbLeftRider.SelectedIndex + 2;
                    }
                    else
                    {
                        cbLeftRider.SelectedIndex = -1;
                    }
                }
            }
            else
            {
                tbRightReaction.Text = "";
                tbRightReaction.Refresh();
                tbRightRaw.Text = "";
                tbRightRaw.Refresh();
                cbRightCones.SelectedIndex = 0;
                if (m_nRaceType == RaceTypes.SingleLane) // single lane
                {
                    if (cbRigthRider.Items.Count > cbRigthRider.SelectedIndex + 1)
                    {
                        cbRigthRider.SelectedIndex = cbRigthRider.SelectedIndex + 1;
                    }
                    else
                    {
                        cbRigthRider.SelectedIndex = -1;
                    }
                }
                else
                {
                    if (cbRigthRider.Items.Count > cbRigthRider.SelectedIndex + 2)
                    {
                        cbRigthRider.SelectedIndex = cbRigthRider.SelectedIndex + 2;
                    }
                    else
                    {
                        cbRigthRider.SelectedIndex = -1;
                    }
                }
            }
        }

        private void cbRound_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbRound.SelectedIndex != -1)
            {
                System.Globalization.CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                cbLeftRider.Items.Clear();
                cbRigthRider.Items.Clear();
                Mac.Excel9.Interop._Worksheet WS = null;

                try{
                    WS = (Mac.Excel9.Interop._Worksheet)theWorkbook.Worksheets.get_Item(cbWorkSheet.SelectedIndex + 1);
                }
                catch (Exception ex)
                {
                    MacMessageBox MMB = new MacMessageBox("Oops! We were not able to read the excelsheet that you previously selected. Did you kill it? Please re-open the spread sheet. " + ex.Message);
                    MMB.StartPosition = FormStartPosition.CenterParent;
                    MMB.ShowDialog();
                }

                Mac.Excel9.Interop.Range range = WS.get_Range("B1", "B99");
                System.Array myvalues = (System.Array)range.Cells.Value2;
                string[] strArray = ConvertToStringArray(myvalues);

                string strStartRound = "";
                int nStartValue = 1;
                Boolean bolFinal = false;
                String strSelRound = cbRound.SelectedItem.ToString();
                String strAddon = "";

                foreach (String cellValue in strArray)
                {
                    nStartValue++;
                    if (strSelRound == "Final & consi")
                    {
                        if (cellValue.ToLower().Contains("cons"))
                        {
                            strStartRound = "B" + nStartValue.ToString();
                            strAddon = " - (consi)";
                        }
                    }
                    else
                    {
                        if (cellValue.IndexOf(strSelRound) != -1)
                        {
                            strStartRound = "B" + nStartValue.ToString();
                        }
                    }

                    if (strStartRound != "")
                    {
                        string strCellRiderName = "B" + nStartValue.ToString();
                        Mac.Excel9.Interop.Range range0 = WS.get_Range(strCellRiderName, strCellRiderName);
                        if (range0.Value2 == null)
                        {
                            if (strSelRound == "Final & consi" && !bolFinal)
                            {
                                bolFinal = true;
                                nStartValue = nStartValue + 3;
                                strCellRiderName = "B" + nStartValue.ToString();
                                range0 = WS.get_Range(strCellRiderName, strCellRiderName);
                                strAddon = " - (final)";
                            }
                            else
                            {
                                break;
                            }
                        }

                        string strRiderName = range0.Cells.Value2.ToString();
                        cbLeftRider.Items.Add(strRiderName + strAddon);
                        cbRigthRider.Items.Add(strRiderName + strAddon);
                    }
                }
                Thread.CurrentThread.CurrentCulture = oldCI;
                tabControl1.SelectedIndex = 1;
            }
        }

        private void bnRefreshList_Click(object sender, EventArgs e)
        {
            if (m_nRaceType == RaceTypes.SingleLane || m_nRaceType == RaceTypes.Qualification)
            {
                Mac.Excel9.Interop._Worksheet WS = null;
                System.Globalization.CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                WS = (Mac.Excel9.Interop._Worksheet)theWorkbook.Worksheets.get_Item(cbWorkSheet.SelectedIndex + 1);
                WS.Activate();
                FillRiderDropDownsSingle(WS);
                Thread.CurrentThread.CurrentCulture = oldCI;
            }
        }

        private int getRoundId()
        {
            if (cbRound.SelectedItem == null)
            {
                return 0;
            }
            else
            {
                if (cbRound.SelectedItem.ToString().Contains("Final"))
                {
                    return 6;
                }
                if (cbRound.SelectedItem.ToString().Contains("4"))
                {
                    return 5;
                }
                if (cbRound.SelectedItem.ToString().Contains("8"))
                {
                    return 4;
                }
                if (cbRound.SelectedItem.ToString().Contains("16"))
                {
                    return 3;
                }
                if (cbRound.SelectedItem.ToString().Contains("32"))
                {
                    return 2;
                }
                if (cbRound.SelectedItem.ToString().Contains("64"))
                {
                    return 1;
                }
            }
            return 0;
        }

        private int getClassId()
        {
            if (cbWorkSheet.SelectedItem.ToString().ToLower().Contains("am"))
            {
                return 1;
            }
            if (cbWorkSheet.SelectedItem.ToString().ToLower().Contains("pro"))
            {
                return 2;
            }
            if (cbWorkSheet.SelectedItem.ToString().ToLower().Contains("wo"))
            {
                return 3;
            }
            if (cbWorkSheet.SelectedItem.ToString().ToLower().Contains("jr") || cbWorkSheet.SelectedItem.ToString().ToLower().Contains("junior"))
            {
                return 4;
            }
            if (cbWorkSheet.SelectedItem.ToString().ToLower().Contains("mas"))
            {
                return 6;
            }
            // open
            return 5;
        }
        
        private void submitRunToWeb(String name, String start, String rawtime, String cones, String lane)
        {
            String isDQ = "0";

            if (rawtime.ToLower().Contains("dq"))
            {
                isDQ = "1";
            }
            if (rawtime.Length == 0)
            {
                rawtime = "0.00";
            }


            String URL = "";
            try
            {
                URL = "http://www.worldcupranking.com/live/admin/reportRun.asp?";

                String strRealName = "";
                int addRound = 0;

                if (name.EndsWith(" - (final)"))
                {
                    strRealName = name.Substring(0, name.Length - 10);
                    addRound++;
                }
                if (name.EndsWith(" - (consi)"))
                {
                    strRealName = name.Substring(0, name.Length - 10);
                }

                URL += "ISSAId=" + tbLiveId.Text;
                URL += "&eventId=" + tbLiveEventId.Text;
                URL += "&classId=" + getClassId();
                URL += "&racer=" + URLEncode(name);
                URL += "&round=" + (getRoundId()+addRound);
                URL += "&start=" + start.Replace(",", ".");
                URL += "&time=" + rawtime.Replace(",", ".");
                URL += "&cones=" + cones;
                URL += "&isDQ=" + isDQ;
                URL += "&lane=" + lane;

                for (int t = 0; t < 5; t++)
                {
                    bool go = false;

                    try
                    {
                        WebRequest request = WebRequest.Create(URL);
                        // If required by the server, set the credentials.
                        request.Credentials = CredentialCache.DefaultCredentials;
                        // Get the response.
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        response.Close();
                    }
                    catch (Exception ex)
                    {
                        String err = ex.Message;
                        go = true;
                    }

                    if (!go)
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MacMessageBox MMB = new MacMessageBox("Oops! The live reporting didn't work. Please check the internetconnection. " + ex.Message);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
            }
        }

        private String URLEncode(String text)
        {
            char[] textArray = text.ToCharArray();
            String URLSafe = "";

            foreach (char tecken in textArray)
            {
                if (tecken > 128)
                {
                    int val = tecken;

                    URLSafe += "%" + val.ToString("X");
                }
                else
                {
                    URLSafe += tecken;
                }
            }

            return URLSafe;
        }

        #endregion

        #region exit
        /// <summary>
        /// Exit the program
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void onExit(object sender, FormClosingEventArgs e)
        {
            StopThread();
            StopDisplayThreadLeft();
            StopDisplayThreadRight();
            System.Globalization.CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            try
            {
                if (tbFileName.Text.Length != 0)
                {
                    theWorkbook.Save();
                    theWorkbook.Close(true, tbFileName.Text, false);
                }
            }catch(Exception ex){
                MacMessageBox MMB = new MacMessageBox("Oops! We couldn't close the Excel spread sheet properly. Perhaps you killed it already? " + ex.Message);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
            }

            objExcel = null;
            Thread.CurrentThread.CurrentCulture = oldCI;
        }

        private void ExcelMate_FormClosing(object sender, FormClosingEventArgs e)
        {
            onExit(null, null);
        }

        #endregion exit        

        #region TrackMateStuff

        private void startThread()
        {
            if (readThread == null)
            {
                m_EventStopThread.Reset();
                readThread = new Thread(new ThreadStart(this.RunThread));
                readThread.Name = "TrackMateReaderThread";
                readThread.Start();
                bnConnect.Text = "Disconnect";
            }
        }

        private void RunThread()
        {
            TrackMateReader TMR = new TrackMateReader(this, m_EventStopThread, m_EventResetTimer);
            try
            {
                TMR.openComPort(cbComPort.Text);
            }
            catch (Exception ex)
            {
                MacMessageBox MMB = new MacMessageBox("Failed to open the com port for trackmate, please close the program and start over. " + ex.Message);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
                return;
            }
            MacMessageBox MMB1 = new MacMessageBox("Connected!");
            MMB1.StartPosition = FormStartPosition.Manual;
            System.Drawing.Point p = this.Location;
            p.X = p.X + this.Width / 2 - MMB1.Width/2;
            p.Y = p.Y + this.Height/ 2 - MMB1.Height/2;
            MMB1.Location = p;
            MMB1.ShowDialog();

            TMR.Read();
        }

        private void StopThread()
        {
            if (readThread != null && readThread.IsAlive)  // thread is active
            {
                // set event "Stop"
                m_EventStopThread.Set();

                // wait when thread  will stop or finish
                while (readThread.IsAlive)
                {
                    Thread.Sleep(100);
                    System.Windows.Forms.Application.DoEvents();
                }
                readThread.Join();
                readThread = null;
            }
        }

        // message ==> the actual time, or other message
        // lane ==> what lane the message refers to. 0 ==> left, 1==> right
        // type ==> reaction or rawtime.             0 ==> reaction, 1 ==> rawtime
        private void AddTrackMateMessage(string message, int lane, int type)
        {
            string strLane = Convert.ToString(lane + 1);

            // reaction time
            if (type == 0)
            {
                if (!cbDiscardReactionTimes.Checked)
                {
                    if (lane == 0 || (m_nRaceType == RaceTypes.SingleLane && cbSingleLanePort.Checked))
                    {
                        tbLeftReaction.Text = message;

                        /* This should be used once the reaction time of "0" in single lane is implemented in the trackmate */
                        if (cbDisplays.Checked && m_bolTrackmateVersion)
                        {
                            if (writeThreadLeft == null)  // thread is active
                            {
                                StartDisplayThreadLeft();
                            }
                        }
                    }
                    else
                    {
                        tbRightReaction.Text = message;
                        /* This should be used once the reaction time of "0" in single lane is implemented in the trackmate */
                        if (cbDisplays.Checked && m_bolTrackmateVersion)
                        {
                            if (writeThreadRight == null)  // thread is active
                            {
                                StartDisplayThreadRight();
                            }
                        }
                    }
                }
            }
            // rawtime
            else
            {
                if (lane == 0 || (m_nRaceType == RaceTypes.SingleLane && cbSingleLanePort.Checked))
                {
                    tbLeftRaw.Text = message;
                }
                else
                {
                    tbRightRaw.Text = message;
                }
            }
            this.WriteToLogfile(message + " lane: " + lane.ToString());
        }

        private string GetTimeString(string tid)
        {
            string hour = Convert.ToString(DateTime.Now.Hour);
            string minute = Convert.ToString(DateTime.Now.Minute);
            string second = Convert.ToString(DateTime.Now.Second + tid.Substring(0, tid.IndexOf(".")));

            if (hour.Length == 1)
            {
                hour = "0" + hour;
            }
            if (minute.Length == 1)
            {
                minute = "0" + minute;
            }
            if (second.Length == 1)
            {
                second = "0" + second;
            }
            return hour + ":" + minute + ":" + second;
        }

        private void bnReset_Click(object sender, EventArgs e)
        {
            DialogResult RS = DialogResult.No;
            if (tbLeftRaw.Text.Length != 0 || tbRightRaw.Text.Length != 0)
            {
                RS = MessageBox.Show("It seems like you haven't saved the latest data. Would you like to save first?", "Forgot to save?", MessageBoxButtons.YesNo);
            }
            if (RS == DialogResult.No)
            {
                if (readThread != null && readThread.IsAlive)  // thread is active
                {
                    // set event "ResetTimer"
                    m_EventResetTimer.Set();
                    Thread.Sleep(1000);
                    m_EventResetTimer.Reset();

                    tbLeftRaw.Text = "";
                    tbRightRaw.Text = "";
                    tbLeftReaction.Text = "";
                    tbRightReaction.Text = "";

                    /*
                     * This should be commented out once we get a reaction of "0" from the trackmate in single lane
                     * As for now the time starts when the timer is restarted.*/
                    if (cbDisplays.Checked && !m_bolTrackmateVersion)
                    {
                        if (writeThreadLeft == null)  // thread is active
                        {
                            StartDisplayThreadLeft();
                        }
                        if (writeThreadRight == null)  // thread is active
                        {
                            StartDisplayThreadRight();
                        }
                    }
                }
                else
                {
                    MacMessageBox MMB = new MacMessageBox("I don't think you got a connection to the Trackmate.");
                    MMB.StartPosition = FormStartPosition.CenterParent;
                    MMB.ShowDialog();
                }
            }
        }

        private void bnConnect_Click_1(object sender, EventArgs e)
        {
            if (bnConnect.Text == "Connect!")
            {
                startThread();
            }
            else
            {
                StopThread();
                bnConnect.Text = "Connect!";
            }
        }

        #endregion

        #region SlalomDisplayStuff

        private void StartDisplayThreadLeft()
        {
            m_EventStopWriteToDisplayLeft.Reset();
            writeThreadLeft = new Thread(new ThreadStart(this.RunDisplayThreadLeft));
            writeThreadLeft.Name = "TrackMateDisplayWriterLeft";
            writeThreadLeft.Start();
        }

        private void StartDisplayThreadRight()
        {
            m_EventStopWriteToDisplayRight.Reset();
            writeThreadRight = new Thread(new ThreadStart(this.RunDisplayThreadRight));
            writeThreadRight.Name = "TrackMateDisplayWriterRight";
            writeThreadRight.Start();
        }

        private void RunDisplayThreadLeft()
        {
            DisplayWriter DWL = new DisplayWriter(this, m_EventStopWriteToDisplayLeft);
            try
            {
                DWL.openComPort(cbDisplayPortLeft.Text);
            }
            catch (Exception ex)
            {
                MacMessageBox MMB = new MacMessageBox("Failed to open the com port for the left display, please close the program and start over. " + ex.Message);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
                return;
            }
            if (m_bolDisplayZeroReaction)
            {
                DWL.WriteReactionToDisplay("0.000");
            }
            else
            {
                DWL.WriteReactionToDisplay(tbLeftReaction.Text);
            }
        }

        private void RunDisplayThreadRight()
        {
            DisplayWriterRight DWR = new DisplayWriterRight(this, m_EventStopWriteToDisplayRight);
            try
            {
                DWR.openComPort(cbDisplayPortRight.Text);
            }
            catch (Exception ex)
            {
                MacMessageBox MMB = new MacMessageBox("Failed to open the com port for the left display, please close the program and start over. " + ex.Message);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
                return;
            }
            if (m_bolDisplayZeroReaction)
            {
                DWR.WriteReactionToDisplay("0.000");
            }
            else
            {
                DWR.WriteReactionToDisplay(tbRightReaction.Text);
            }
        }

        private void StopDisplayThreadLeft()
        {
            if (writeThreadLeft != null && writeThreadLeft.IsAlive)  // thread is active
            {
                try
                {
                    // set event "Stop"
                    m_EventStopWriteToDisplayLeft.Set();

                    // wait when thread  will stop or finish
                    while (writeThreadLeft.IsAlive)
                    {
                        Thread.Sleep(100);
                        System.Windows.Forms.Application.DoEvents();
                    }
                    writeThreadLeft.Join();
                    writeThreadLeft = null;
                }
                catch (Exception ex)
                {
                    String err = ex.Message;
                }
            }
        }

        private void StopDisplayThreadRight()
        {
            if (writeThreadRight != null && writeThreadRight.IsAlive)  // thread is active
            {
                try
                {
                    // set event "Stop"
                    m_EventStopWriteToDisplayRight.Set();

                    // wait when thread  will stop or finish
                    while (writeThreadRight.IsAlive)
                    {
                        Thread.Sleep(100);
                        System.Windows.Forms.Application.DoEvents();
                    }
                    writeThreadRight.Join();
                    writeThreadRight = null;
                }
                catch (Exception ex)
                {
                    String err = ex.Message;
                }
            }
        }

        private void StartFinalTimeDisplayThreadLeft()
        {
            m_EventStopWriteToDisplayLeft.Reset();
            writeThreadLeft = new Thread(new ThreadStart(this.RunFinalTimeDisplayThreadLeft));
            writeThreadLeft.Name = "TrackMateDisplayWriterLeft";
            writeThreadLeft.Start();
        }

        private void StartFinalTimeDisplayThreadRight()
        {
            m_EventStopWriteToDisplayRight.Reset();
            writeThreadRight = new Thread(new ThreadStart(this.RunFinalTimeDisplayThreadRight));
            writeThreadRight.Name = "TrackMateDisplayWriterRight";
            writeThreadRight.Start();
        }

        private void RunFinalTimeDisplayThreadLeft()
        {
            DisplayWriter DWL = new DisplayWriter(this, m_EventStopWriteToDisplayLeft);
            DWL.openComPort(cbDisplayPortLeft.Text);
            //DWL.openComPort("COM2");
            DWL.WriteFinalTimeToDisplay(tbLeftRaw.Text);
        }

        private void RunFinalTimeDisplayThreadRight()
        {
            DisplayWriterRight DWR = new DisplayWriterRight(this, m_EventStopWriteToDisplayRight);
            DWR.openComPort(cbDisplayPortRight.Text);
            //DWR.openComPort("COM2");
            DWR.WriteFinalTimeToDisplay(tbRightRaw.Text);
        }

        private void ShowFinalTimeOnDisplayLeft(object sender, EventArgs e)
        {
            if (cbDisplays.Checked && writeThreadLeft != null && writeThreadLeft.IsAlive)  // thread is active
            {
                StopDisplayThreadLeft();
                StartFinalTimeDisplayThreadLeft();
                StopDisplayThreadLeft();
            }
        }

        private void ShowFinalTimeOnDisplayRight(object sender, EventArgs e)
        {
            if (cbDisplays.Checked && writeThreadRight != null && writeThreadRight.IsAlive)  // thread is active
            {
                StopDisplayThreadRight();
                StartFinalTimeDisplayThreadRight();
                StopDisplayThreadRight();
            }
        }

        private void bnTest_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex != 0 && cbDisplays.Checked && (cbComPort.Text == cbDisplayPortLeft.Text || cbComPort.Text == cbDisplayPortRight.Text || cbDisplayPortLeft.Text == cbDisplayPortRight.Text))
            {
                MacMessageBox MMB = new MacMessageBox("You need to select different comports for each connection if you're using external displays.");
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
            }
            else
            {
                MacMessageBox MMB = new MacMessageBox("Connected!");
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
            }
        }

        private void cbDisplays_CheckedChanged(object sender, EventArgs e)
        {
            if (cbDisplays.Checked)
            {
                if (MessageBox.Show("If you have Trackmate v6.5 or higher click 'Yes' otherwise click 'No' and the times on the displays will start rolling at the fourth beep.", "Trackmate version", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    m_bolTrackmateVersion = true;
                }
                else
                {
                    m_bolTrackmateVersion = false;
                }
                if (MessageBox.Show("If this race uses individual starts (no reaction time) click yes, otherwise click no", "Use reaction", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    m_bolDisplayZeroReaction = true;
                }
                else
                {
                    m_bolDisplayZeroReaction = false;
                }
                label10.Visible = true;
                label11.Visible = true;
                label12.Visible = true;
                cbDisplayPortLeft.Visible = true;
                cbDisplayPortRight.Visible = true;
                bnConnectDisplays.Visible = true;
            }
            else
            {
                label10.Visible = false;
                label11.Visible = false;
                label12.Visible = false;
                cbDisplayPortLeft.Visible = false;
                cbDisplayPortRight.Visible = false;
                bnConnectDisplays.Visible = false;
            }
        }


        private void ShowFinalTimeOnDisplayLeft()
        {
            ShowFinalTimeOnDisplayLeft(null, null);
        }

        private void ShowFinalTimeOnDisplayRight()
        {
            ShowFinalTimeOnDisplayRight(null, null);
        }

        #endregion

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.cbPreviousData.Parent = this.tabRace;
            this.bnRefreshList.Parent = this.tabRace;
            if (cbLayOut.Checked) // old layout
            {
                    System.Drawing.Point pt = cbLeftRider.Location;
                    pt.X = 46;
                    pt.Y = 14;
                    cbLeftRider.Location = pt;
                    pt.X = 408;
                    pt.Y = 14;
                    tbLeftReaction.Location = pt;
                    pt.X = 489;
                    pt.Y = 14;
                    tbLeftRaw.Location = pt;
                    pt.X = 583;
                    pt.Y = 13;
                    cbLeftCones.Location = pt;
                    pt.X = 658;
                    pt.Y = 13;
                    bnSaveLeft.Location = pt;

                    pt.X = 46;
                    pt.Y = 58;
                    cbRigthRider.Location = pt;
                    pt.X = 408;
                    pt.Y = 58;
                    tbRightReaction.Location = pt;
                    pt.X = 489;
                    pt.Y = 58;
                    tbRightRaw.Location = pt;
                    pt.X = 583;
                    pt.Y = 58;
                    cbRightCones.Location = pt;
                    pt.X = 658;
                    pt.Y = 58;
                    bnSaveRight.Location = pt;

                    pt.X = 770;
                    pt.Y = 13;
                    bnReset.Location = pt;

                    pt.X = 770;
                    pt.Y = 51;
                    bnRefreshList.Location = pt;

                    pt.X = 407;
                    pt.Y = 1;
                    label7.Location = pt;
                    pt.X = 487;
                    pt.Y = 1;
                    label8.Location = pt;
                    pt.X = 578;
                    pt.Y = 0;
                    label9.Location = pt;

                    pt.X = 407;
                    pt.Y = 45;
                    label14.Location = pt;
                    pt.X = 487;
                    pt.Y = 45;
                    label13.Location = pt;
                    pt.X = 578;
                    pt.Y = 45;
                    label6.Location = pt;

                    label4.Visible = true;
                    label3.Visible = true;
                    label14.Visible = false;
                    label13.Visible = false;
                    label6.Visible = false;

                    if (m_nRaceType == RaceTypes.SingleLane)
                    {
                        label4.Visible = false;
                        label14.Visible = false;
                        label13.Visible = false;
                        label6.Visible = false;

                        cbRigthRider.Visible = false;
                        tbRightReaction.Visible = false;
                        tbRightRaw.Visible = false;
                        cbRightCones.Visible = false;

                        bnSaveRight.Visible = false;

                        label3.Text = "Rider";

                        pt = cbPreviousData.Location;
                        pt.X = 660;
                        pt.Y = 54;
                        cbPreviousData.Location = pt;

                        Size sz = tbPrevData.Size;
                        sz.Height = 155;
                        tbPrevData.Size = sz;
                        pt.X = 44;
                        pt.Y = 52;
                        tbPrevData.Location = pt;
                    }
                    else
                    {
                        label4.Visible = true;
                        label14.Visible = true;
                        label13.Visible = true;
                        label6.Visible = true;

                        cbRigthRider.Visible = true;
                        tbRightReaction.Visible = true;
                        tbRightRaw.Visible = true;
                        cbRightCones.Visible = true;
                        label3.Text = "Left";

                        pt = cbPreviousData.Location;
                        pt.X = 772;
                        pt.Y = 75;
                        cbPreviousData.Location = pt;

                        Size sz = tbPrevData.Size;
                        sz.Height = 115;
                        tbPrevData.Size = sz;
                        pt.X = 44;
                        pt.Y = 92;
                        tbPrevData.Location = pt;
                    }

            }
            else // new layout
            {
                if (m_nRaceType != RaceTypes.SingleLane)
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
                    System.Drawing.Point pt = cbLeftRider.Location;
                    pt.X = 46;
                    pt.Y = 14;
                    cbLeftRider.Location = pt;
                    pt.X = 408;
                    pt.Y = 14;
                    tbLeftReaction.Location = pt;
                    pt.X = 489;
                    pt.Y = 14;
                    tbLeftRaw.Location = pt;
                    pt.X = 583;
                    pt.Y = 13;
                    cbLeftCones.Location = pt;
                    pt.X = 658;
                    pt.Y = 13;
                    bnSaveLeft.Location = pt;

                    label4.Visible = false;
                    label14.Visible = false;
                    label13.Visible = false;
                    label6.Visible = false;

                    cbRigthRider.Visible = false;
                    tbRightReaction.Visible = false;
                    tbRightRaw.Visible = false;
                    cbRightCones.Visible = false;

                    bnSaveRight.Visible = false;

                    label3.Text = "Rider";

                    pt = cbPreviousData.Location;
                    pt.X = 660;
                    pt.Y = 54;
                    cbPreviousData.Location = pt;

                    Size sz = tbPrevData.Size;
                    sz.Height = 155;
                    tbPrevData.Size = sz;
                    pt.X = 44;
                    pt.Y = 52;
                    tbPrevData.Location = pt;
                }
            }
            if (this.cbColor.Checked)
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
            if (tabControl1.SelectedIndex != 0 && cbDisplays.Checked && (cbComPort.Text == cbDisplayPortLeft.Text || cbComPort.Text == cbDisplayPortRight.Text || cbDisplayPortLeft.Text == cbDisplayPortRight.Text))
            {
                MacMessageBox MMB = new MacMessageBox("You need to select different comports for each connection if you're using external displays.");
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
                tabControl1.SelectedIndex = 0;
            }
            if (tabControl1.SelectedIndex != 0 && (tbFileName.Text.Length == 0 || theWorkbook == null))
            {
                MacMessageBox MMB = new MacMessageBox("Please select an Excel file.");
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
                tabControl1.SelectedIndex = 0;
            }
            if (tabControl1.SelectedIndex != 0 && m_nRaceType == RaceTypes.NotSet)
            {
                MacMessageBox MMB = new MacMessageBox("Please select worksheet.");
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
                tabControl1.SelectedIndex = 0;
            }
            if (tabControl1.SelectedIndex != 0 && m_nRaceType == RaceTypes.Elimination && cbRound.SelectedIndex == -1)
            {
                MacMessageBox MMB = new MacMessageBox("Please select which round of the race we are running.");
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
                tabControl1.SelectedIndex = 0;
            }

            if (cbDiscardReactionTimes.Checked)
            {
                tbRightReaction.BackColor = Color.Gray;
                tbLeftReaction.BackColor = Color.Gray;            
            }
            else
            {
                tbRightReaction.BackColor = Color.White;
                tbLeftReaction.BackColor = Color.White;
            }

            if (tabControl1.SelectedIndex == 0)
            {
                    this.Height = 171;
            }
            else
            {
                if (m_nRaceType == RaceTypes.SingleLane){
                    this.Height = 131;
                }else{
                    this.Height = 151;
                }
            }

        }

        // Kanske att vi ska titta på den här
        private void cbPreviousData_CheckedChanged(object sender, EventArgs e)
        {
            if (m_nRaceType == RaceTypes.SingleLane)
            {
                if (cbPreviousData.Checked)
                {
                    tbPrevData.Visible = true;
                    this.Height = 270;
                }
                else
                {
                    tbPrevData.Visible = false;
                    this.Height = 131;
                }
            }
            else
            {
                if (cbPreviousData.Checked)
                {
                    tbPrevData.Visible = true;
                    this.Height = 270;
                }
                else
                {
                    tbPrevData.Visible = false;
                    this.Height = 151;
                }
            }
        }

        private void cbLiveReport_CheckedChanged(object sender, EventArgs e)
        {
            if (cbLiveReport.Checked)
            {
                label15.Visible = true;
                tbLiveId.Visible = true;
                tbLiveId.Enabled = true;
                label16.Visible = true;
                tbLiveEventId.Visible = true;
                tbLiveEventId.Enabled = true;
                bnCheckId.Visible = true;
            }
            else
            {
                label15.Visible = false;
                tbLiveId.Visible = false;
                tbLiveId.Enabled = false;
                label16.Visible = false;
                tbLiveEventId.Visible = false;
                tbLiveEventId.Enabled = false;
                bnCheckId.Visible = false;
            }
        }

        private void LeftConesChanged(object sender, EventArgs e)
        {
            if (cbLeftCones.SelectedItem.ToString() == "DQ")
            {
                MacMessageBox MMB = new MacMessageBox("Should this race be a DQ?", MessageBoxButtons.OKCancel);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
                if (MMB.DialogResult == DialogResult.OK)
                {
                    cbLeftCones.SelectedIndex = 0;
                    tbLeftRaw.Text = "DQ";
                }
            }
        }

        private void RightConesChanged(object sender, EventArgs e)
        {
            if (cbRightCones.SelectedItem.ToString() == "DQ")
            {
                MacMessageBox MMB = new MacMessageBox("Should this race be a DQ?", MessageBoxButtons.OKCancel);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();
                if (MMB.DialogResult == DialogResult.OK)
                {
                    cbRightCones.SelectedIndex = 0;
                    tbRightRaw.Text = "DQ";
                }
            }
        }

        private void bnCheckId_Click(object sender, EventArgs e)
        {
            MacMessageBox MMB = new MacMessageBox("This check can take 30 seconds, please be patient", MessageBoxButtons.OK);
            MMB.StartPosition = FormStartPosition.CenterParent;
            MMB.ShowDialog();

            try
            {
                String URL = "";
                if (tbLiveId.Text == "666")
                {
                    if (m_strReportSiteSelection == "")
                    {
                        if (MessageBox.Show("Klicka 'ja' för 161 eller 'nej' för att rapportera till wcrank", "välj DB", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            m_strReportSiteSelection = "http://www.ettsexett.com/live/admin/checkEvent.asp?eventId=" + tbLiveEventId.Text + "&ISSAId=" + tbLiveId.Text;
                        }
                        else
                        {
                            m_strReportSiteSelection = "http://www.worldcupranking.com/live/admin/checkEvent.asp?eventId=" + tbLiveEventId.Text + "&ISSAId=" + tbLiveId.Text;
                        }
                    }
                    URL = m_strReportSiteSelection;
                }
                else
                {
                    URL = "http://www.worldcupranking.com/live/admin/checkEvent.asp?eventId=" + tbLiveEventId.Text +"&ISSAId="+ tbLiveId.Text;
                }

                WebRequest request = WebRequest.Create(URL);

                // If required by the server, set the credentials.
                request.Credentials = CredentialCache.DefaultCredentials;

                // Get the response.
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                // HÄMTA HEM DET HELA HÄR
                Encoding enc = System.Text.Encoding.GetEncoding(1252);
                StreamReader loResponseStream = new StreamReader(response.GetResponseStream(), enc);
                String strResponse = loResponseStream.ReadToEnd();

                MMB = new MacMessageBox(strResponse, MessageBoxButtons.OK);
                MMB.StartPosition = FormStartPosition.CenterParent;
                MMB.ShowDialog();

                loResponseStream.Close();
                response.Close();
            }
            catch (Exception ex)
            {
                String err = ex.Message;
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


        private void cbColors_CheckedChanged(object sender, EventArgs e)
        {
            if (this.cbColor.Checked)
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

        private Color GetColor(string name)
        {
            switch (name)
            {
                case "Red":
                    return Color.Red;

                case "Green":
                    return Color.Green;

                case "Orange":
                    return Color.Orange;

                case "Blue":
                    return Color.Blue;
            }
            return Color.White;
        }

        private void WriteToLogfile(string message)
        {
            try
            {
                if (this.cbLog2File.Checked)
                {
                    if (!File.Exists(this.m_strLogFile))
                    {
                        File.Create(this.m_strLogFile);
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

    }
}