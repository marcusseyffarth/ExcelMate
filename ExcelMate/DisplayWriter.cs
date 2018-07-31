using System;
using System.Collections.Generic;
using System.Text;
using System.IO.Ports;
using System.Threading;

namespace ExcelMate
{
    public class DisplayWriter
    {
        mExcelMate myTimer;
        static SerialPort seriePortDisplay;
        private ManualResetEvent m_EventStopWriteToDisplay;
        static bool WRITE;
        private long doh = 0;

        public DisplayWriter(mExcelMate timer, ManualResetEvent rES)
        {
            myTimer = timer;
            m_EventStopWriteToDisplay = rES;
        }

        public void openComPort(string portName)
        {
            // if we click twice, the port should not be opened twice.
            // this is where we actually open the port and create a 
            // new thread within the application, so that the app can 
            // listen on the port and still respond to other activities

            seriePortDisplay = new SerialPort();

            // set some appropriate properties.
            seriePortDisplay.PortName = portName;
            seriePortDisplay.BaudRate = 9600;
            seriePortDisplay.Parity = System.IO.Ports.Parity.None;
            seriePortDisplay.DataBits = 8;
            seriePortDisplay.StopBits = System.IO.Ports.StopBits.One;
            seriePortDisplay.Handshake = System.IO.Ports.Handshake.None;

            // Set the read/write timeouts
            seriePortDisplay.ReadTimeout = 500;
            seriePortDisplay.WriteTimeout = 500;

            if (!seriePortDisplay.IsOpen)
            {
                // boolean variable to be able to exit the reading loop below
                WRITE = true;

                // open the port
                seriePortDisplay.Open();
            }
        }

        private void CloseAndClean()
        {
            if (seriePortDisplay.IsOpen)
            {
                seriePortDisplay.Close();
            }
        }

        private void WriteToDisplay()
        {
            // until this boolean is false - read on the serial port
            while (WRITE)
            {
                try
                {
                    // check if thread is cancelled
                    if (m_EventStopWriteToDisplay.WaitOne(0, true))
                    {
                        // clean-up operations may be placed here
                        // ...
                        // inform main thread that this thread stopped
                        WRITE = false;
                    }

                    // write to port
                    TimeSpan TS = new TimeSpan(DateTime.Now.Ticks - doh);
                    String strMillis = TS.Milliseconds.ToString();
                    if (strMillis.Length == 3){
                        strMillis = strMillis + " ";
                    }
                    if (strMillis.Length == 2){
                        strMillis = strMillis + "0 ";
                    }
                    if (strMillis.Length == 1){
                        strMillis = strMillis + "00 ";
                    }
                    if (strMillis.Length == 0){
                        strMillis = strMillis + "000 ";
                    }
                    String strTime = TS.Seconds.ToString() + "." + strMillis;
                    seriePortDisplay.WriteLine(strTime);

                    Thread.Sleep(100);
                }
                catch (TimeoutException) { }
            }
            CloseAndClean();
        }

        public void WriteFinalTimeToDisplay(string time)
        {
            seriePortDisplay.WriteLine(time+" ");
            CloseAndClean();
        }

        public void WriteReactionToDisplay(string time)
        {
            doh = DateTime.Now.Ticks;
            seriePortDisplay.WriteLine(time +" ");
            Thread.Sleep(3000);
            WriteToDisplay();
        }

    }
}
