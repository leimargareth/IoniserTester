using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Windows.Forms;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;

namespace IoniserTester
{
    public partial class Form2 : Form
    {
        public static SerialPort spDLY = new SerialPort();
        public const int RxDataSize = 140;
        public static byte[] rx_data = new byte[RxDataSize];
        public static int rx_len = 0;
        public static bool reading_OL = false;
        public static bool reading_valid = false;
        public static DateTime dtLastRxValidReading = DateTime.Now;
        public static string reading_string = "";
        public static double reading = 0.0;
        public static double ion_multiply_factor = 0.01;

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
        public static void OpenPort(string portName)
        {
            if (!spDLY.IsOpen)
            {
                spDLY.PortName = portName;
                spDLY.BaudRate = 2400;
                spDLY.DataBits = 8;
                spDLY.Parity = Parity.None;
                spDLY.Open();
            }
            var taskDLYMainLoop = new Thread(MainLoop);
            taskDLYMainLoop.IsBackground = true;
            taskDLYMainLoop.Start();
        }

        public static void Set_Ion_Mutliply_Factor(double factor)
        {
            ion_multiply_factor = factor;
        }

        public static int Read_Ions_Reading(out double ion_reading, out string ion_reading_string, out bool over_flow)
        {
            DateTime dtStart = DateTime.Now;
            reading_valid = false;
            over_flow = false;
            ion_reading = 0;
            ion_reading_string = "";
            do
            {
                Thread.Sleep(10);
                if (reading_valid)
                {
                    ion_reading = reading;
                    ion_reading_string = reading_string;
                    over_flow = reading_OL;
                    return 0;
                }
            } while ((DateTime.Now - dtStart).TotalSeconds < 1);
            return (-1);
        }

        public static void MainLoop()
        {
            for (; ; )
            {
                Thread.Sleep(100);
                if (!spDLY.IsOpen) continue;
                TimeSpan ts = DateTime.Now - dtLastRxValidReading;
                if (ts.TotalSeconds > 2)
                {
                    reading_valid = false;
                    reading_OL = false;
                }

                int len = spDLY.BytesToRead;
                if (len > 0)
                {
                    Thread.Sleep(30);
                    byte[] rx_buf = new byte[141];
                    if (len > RxDataSize) len = RxDataSize;
                    spDLY.Read(rx_buf, 0, len);
                    for (int i = 0; i < len; i++) rx_data[i] = rx_buf[i];
                    rx_len = len;
                    if (rx_len >= 14)
                    {
                        if ((rx_data[0] == 0x15) && (rx_data[13] == 0xE4))
                        {
                            string sSign = "";
                            if ((rx_buf[1] & 0x07) == 0x00)
                            {
                                reading_OL = true;
                            }
                            else
                            {
                                reading_OL = false;
                            }
                            {
                                if ((rx_data[1] & 0x08) == 0x08) sSign = "-";
                                byte[] data = new byte[4];
                                string[] sData = new string[4];
                                bool invalid_format = false;
                                for (int i = 0; i < 4; i++)
                                {
                                    byte msb = (byte)(rx_data[i * 2 + 1] & 0x07);
                                    byte lsb = (byte)(rx_data[i * 2 + 2] & 0x0F);
                                    data[i] = (byte)(msb * 16 + lsb);
                                    string s = "0";
                                    switch (data[i])
                                    {
                                        case 0x7D: s = "0"; break;
                                        case 0x05: s = "1"; break;
                                        case 0x5B: s = "2"; break;
                                        case 0x1F: s = "3"; break;
                                        case 0x27: s = "4"; break;
                                        case 0x3E: s = "5"; break;
                                        case 0x7E: s = "6"; break;
                                        case 0x15: s = "7"; break;
                                        case 0x7F: s = "8"; break;
                                        case 0x3F: s = "9"; break;
                                        default: invalid_format = true; break;
                                    }
                                    sData[i] = s;
                                }
                                if (!invalid_format)
                                {
                                    reading_string = sSign + sData[0] + sData[1] + sData[2] + sData[3];

                                    try
                                    {
                                        reading = ion_multiply_factor * double.Parse(reading_string);
                                        dtLastRxValidReading = DateTime.Now;
                                        reading_valid = true;
                                        Console.WriteLine("Value: " + reading + "  DATE:  " + dtLastRxValidReading + "  STATUS: " + reading_valid);

                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                            }
                        }
                    }
                }
            }
        }


        private void buttonOpen_Click(object sender, EventArgs e)
        {
            if (!spDLY.IsOpen)
            {
                spDLY.PortName = "COM4";
                spDLY.BaudRate = 2400;
                spDLY.DataBits = 8;
                spDLY.Parity = Parity.None;

                spDLY.Open();
            }
            var taskDLYMainLoop = new Thread(MainLoop);
            taskDLYMainLoop.IsBackground = true;
            taskDLYMainLoop.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            this.textBoxReceivedData.BeginInvoke(
            (Action)(() =>
            {
                this.textBoxReceivedData.AppendText(Environment.NewLine + "Value: " + reading + "   DATE: " + dtLastRxValidReading + "STATUS:  " + reading_valid);

            }));
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            OpenPort("COM4");
            MessageBox.Show("Hi", "");
        }
    }
}
