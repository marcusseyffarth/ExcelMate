using System;
using System.Collections.Generic;
using System.Text;
using System.IO.Ports;
using System.Threading;

namespace ExcelMate
{
    public class TrackMateReader
    {
        mExcelMate myTimer;
        static SerialPort seriePort;
        private ManualResetEvent m_EventStop;
        private ManualResetEvent m_EventResetTimer;
        static bool READ;

        public TrackMateReader(mExcelMate timer, ManualResetEvent resetEventStop, ManualResetEvent resetTimer)
        {
            myTimer = timer;
            m_EventStop = resetEventStop;
            m_EventResetTimer = resetTimer;
        }

        public void openComPort(string portName)
        {
            // if we click twice, the port should not be opened twice.
            // this is where we actually open the port and create a 
            // new thread within the application, so that the app can 
            // listen on the port and still respond to other activities

            seriePort = new SerialPort();

            // set some appropriate properties.
            seriePort.PortName = portName;
            seriePort.BaudRate = 9600;
            seriePort.Parity = System.IO.Ports.Parity.None;
            seriePort.DataBits = 8;
            seriePort.StopBits = System.IO.Ports.StopBits.One;
            seriePort.Handshake = System.IO.Ports.Handshake.None;

            // Set the read/write timeouts
            seriePort.ReadTimeout = 500;
            seriePort.WriteTimeout = 500;

            if (!seriePort.IsOpen)
            {
                // boolean variable to be able to exit the reading loop below
                READ = true;

                // open the port
                seriePort.Open();
            }
        }

        private void CloseAndClean()
        {
            if (seriePort.IsOpen)
            {
                seriePort.Close();
            }
        }

        public void Read()
        {
            // until this boolean is false - read on the serial port
            while (READ)
            {
                try
                {
                    // check if thread is cancelled
                    if (m_EventStop.WaitOne(0, true))
                    {
                        // clean-up operations may be placed here
                        // ...
                        // inform main thread that this thread stopped
                        READ = false;
                    }

                    // Reset the timer
                    if (m_EventResetTimer.WaitOne(0, true))
                    {
                        // write "R" to the serial port and the timer should restart.
                        seriePort.WriteLine("R");
                    }

                    // read from port
                    string message = seriePort.ReadLine();

                    if (message.EndsWith("\r"))
                    {
                        message = message.Substring(0, message.Length - 1);
                    }

                    // this is for us! String should be of type @nnLxTmmm[CR+LF]
                    if (message.StartsWith("@"))
                    {
                        // some variables
                        string raceType = "";
                        string lane = "";
                        string time = "";
                        string timeDec = "";

                        // get the racetype 00 or 01
                        raceType = message.Substring(2, 1);

                        // get the lane L1 or L2
                        lane = message.Substring(4, 1);

                        // get the time Txxxxx where xxxx is time in milliseconds, perhaps with a starting '-'
                        time = message.Substring(6);

                        // if this is not a negative time, and the length of the string is more than 3 characters
                        if (!time.StartsWith("-"))
                        {
                            // time string contains one digit. 1/1000 of a second
                            if (time.Length == 1)
                            {
                                timeDec = "0.00" + time;
                            }
                            // time string contains two digits. 1/100 of a second
                            if (time.Length == 2)
                            {
                                timeDec = "0.0" + time;
                            }
                            // time string contains three digits. 1/10 of a second
                            if (time.Length == 3)
                            {
                                timeDec = "0." + time;
                            }
                            // one digit to the left of the decimal sign
                            if (time.Length == 4)
                            {
                                timeDec = time.Substring(0, 1) + "." + time.Substring(1);
                            }
                            // two digits to the left of the decimal sign
                            if (time.Length == 5)
                            {
                                timeDec = time.Substring(0, 2) + "." + time.Substring(2);
                            }
                            // three digits to the left of the decimal sign
                            if (time.Length == 6)
                            {
                                timeDec = time.Substring(0, 3) + "." + time.Substring(3);
                            }
                        }

                        // negative case. Pretty much same as above, though you can only start
                        // 9.999 seconds early.
                        if (time.StartsWith("-"))
                        {
                            if (time.Length == 2)
                            {
                                timeDec = "-0.00" + time.Substring(1);
                            }
                            if (time.Length == 3)
                            {
                                timeDec = "-0.0" + time.Substring(1);
                            }
                            if (time.Length == 4)
                            {
                                timeDec = "-0." + time.Substring(1);
                            }
                            if (time.Length == 5)
                            {
                                timeDec = "-" + time.Substring(1, 1) + "." + time.Substring(2);
                            }
                        }

                        // if the racetype is reaction we output that
                        if (raceType == "0")
                        {
                            if (myTimer != null)
                            {
                                myTimer.Invoke(myTimer.mTrackMateCallback, timeDec, Convert.ToInt32(lane), Convert.ToInt32(raceType));
                            }
                        }
                        // rawtime
                        else
                        {
                            // race time
                            if (myTimer != null)
                            {
                                myTimer.Invoke(myTimer.mTrackMateCallback, timeDec, Convert.ToInt32(lane), Convert.ToInt32(raceType));
                            }
                        }
                    }
                }
                catch (TimeoutException) { }
            }
            CloseAndClean();
        }    
    }
}
