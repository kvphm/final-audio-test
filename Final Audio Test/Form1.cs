using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ini;
using System.Drawing.Printing;
//using AIOWDMNet;  //TempDebug
using AudioPrecision.API;
using System.Threading;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Final_Audio_Test
{
    public partial class Form1 : Form
    {
        //   UInt32 Status = 0;
        /*****************************************************************************************/
        /*** Basic data that we will be using to interface with the DIO card:
        /*****************************************************************************************/
        const UInt32 MAX_CARDS = 10;
        public Int32 NumCards = 0;
        public Int32 CardNum = 0; // The index of the one card we will be using 0
        public UInt32 Offset = 0; // This the offset that will be used based on the base address of the card

        public bool RunFlag = false;
        public bool TimerEnabled = false;
        public bool ReportTabSelected = false;

        public byte[] value = new byte[] {
      0,
      0,
      0
    }; // 3 values used for read write
        public int i = 0;

        // This will hold the data for the card installed in the system:
        public struct TCardData
        {
            public bool IsValid;
            public bool IsSelected;
            public UInt32 DeviceID;
            public UInt32 Base; // This a result of findcards and the base address of the card
        };
        public TCardData CardData; // only one struct	

        // DIO offsets +4 each decimal 0,4,8,12,16
        public UInt32[] PortOffsets = new UInt32[5] {
      0x0,
      0x4,
      0x8,
      0x12,
      0x16
    }; // offsets for 5 max groups

        // Interval Timer for GUI update:
        // (Note: A Windows Forms timer designed for single threaded environments not a System Timer 55 ms min res)
        static System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();
        /***************************************************************************************************************/

        private APx500 APx;

        Bitmap bitmap;

        Product DUT = new Product();
        productionDrive Folder = new productionDrive();

        IniFile ini = new IniFile(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Final Audio.ini");

        public Form1()
        {
            InitializeComponent();

            /*  AIOWDM.RelOutPortB(CardNum, Offset + 3, 0x80);
              AIOWDM.RelOutPortB(CardNum, Offset, 0x00);
              AIOWDM.RelOutPortB(CardNum, Offset + 1, 0x00);
              AIOWDM.RelOutPortB(CardNum, Offset + 2, 0x00);  */ //TempDebug

            DUT.productModel = "100X";
            setupRefData();

        }

        void setupRefData()
        {
            try
            {
                DUT.noBlackBoxSetup = Convert.ToInt16(ini.IniReadValue("RefData", "noBlackBoxSetup"));
                DUT.autoShutdown = Convert.ToInt16(ini.IniReadValue("RefData", "autoShutdown"));
                DUT.APautoShutdown = Convert.ToInt16(ini.IniReadValue("RefData", "APautoShutdown"));
                DUT.Debug = Convert.ToInt16(ini.IniReadValue("RefData", "Debug"));
                DUT.freqStart = Convert.ToInt16(ini.IniReadValue("RefData", "freqStart"));
                DUT.freqStop = Convert.ToInt16(ini.IniReadValue("RefData", "freqStop"));
                DUT.ReportFormPN = Convert.ToString(ini.IniReadValue("RefData", "ReportFormPN"));

                Folder.networkDrive = ini.IniReadValue("RefData", "networkDrive");
                Folder.setupPicsFolder = ini.IniReadValue("RefData", "setupPicsFolder");
                Folder.refFileFolder = ini.IniReadValue("RefData", "refFileFolder");
                Folder.testDataFolder = ini.IniReadValue(DUT.productModel, "testDataFolder");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            DUT.resetVariablesToDefault();
        }

        /********************************************************************************************/
        /* To shutdown the system including APx500 4.6 after 30 minutes of no activity was detected */
        /********************************************************************************************/
        private void shutdownTimer_Tick(object sender, EventArgs e)
        {
            if (DUT.APautoShutdown == 1)
            {
                if (APx != null)
                {
                    APx.Exit();
                }
            }

            if (DUT.autoShutdown == 1)
            {
                //   remarkBox.Text += " - Timer Shutdown";                        
                tabControl1.SelectedTab = tabControl1.TabPages["tabPage1"];
                Form1.ActiveForm.Update();
                Thread.Sleep(200);

                SaveData(DUT.FileNameExt);
                shutdownTimer.Stop();
                channelSetup(0x00, 0x00, 0x00);
                Application.Exit();
            }
        }

        void Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.WindowsShutDown)
            {
                tabControl1.SelectedTab = tabControl1.TabPages["tabPage1"];
                Form1.ActiveForm.Update();
                Thread.Sleep(200);
                SaveData(DUT.FileNameExt);
            }

        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            tabControl1.SelectedTab = tabControl1.TabPages["tabPage1"];
            Form1.ActiveForm.Update();

            Thread.Sleep(200);
            //  remarkBox.Text += " - User Shutdown";            
            SaveData(DUT.FileNameExt);

            shutdownTimer.Stop();
            if (DUT.APautoShutdown == 1)
            {
                if (APx != null)
                {
                    APx.Exit();
                }
            }
        }

        void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                //  remarkBox.Text += "ESC key pressed, "; remarkBox.Update();
                //  APx.CancelOperation();
            }
            else if ((e.Modifiers == Keys.Shift && e.KeyCode == Keys.P) && (ReportTabSelected))
            {

                bitmap = new Bitmap(this.Width, this.Height);
                this.DrawToBitmap(bitmap, new Rectangle(0, 0, this.Width, this.Height)); //bitmap is a global variable

                PrintDialog printDlg = new PrintDialog();
                PrintDocument printDoc = printDocument;

                printDocument.OriginAtMargins = true;

                printDlg.Document = printDoc;
                printDlg.AllowSelection = true;
                printDlg.AllowSomePages = true;

                //Call ShowDialog
                if (printDlg.ShowDialog() == DialogResult.OK)
                {
                    printDoc.DefaultPageSettings.Landscape = false;
                    printDoc.DefaultPageSettings.Margins = new Margins(80, 80, 65, 30);
                    printDoc.DefaultPageSettings.Color = true;

                    printDoc.Print();
                }
            }
        }

        private void PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //Print the contents.
            e.Graphics.DrawImage(bitmap, 0, 0);
        }

        private void tab1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"]) //Print Report page
            {
                // MessageBox.Show("Tab 5 Is Selected");
                ReportTabSelected = true;
            }
            else
            {
                ReportTabSelected = false;
            }

            channelSetup(0x00, 0x00, DUT.portCState);
        }

        private void dispose() { }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            this.KeyPreview = true;

            try
            {
                APx = new APx500(APxOperatingMode.SequenceMode);
                APx.Visible = true;
                APx.Maximize();

                if (APx.IsDemoMode)
                {
                    MessageBox.Show("FAILED TO CONNECT! AP IS IN DEMO MODE.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            try
            {
                if (DUT.noBlackBoxSetup == 1)
                {
                    APx.OpenProject(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Final Audio (Test Cart).approjx");
                }
                else
                {
                    APx.OpenProject(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Final Audio (Production).approjx");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " In folder ''" + Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "''", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //     if (!APx.FftSpectrumSignalMonitor.IsDocked) APx.FftSpectrumSignalMonitor.Dock();
            // TEMP DEBUG
            // if (!APx.ScopeSignalMonitor.IsDocked) APx.ScopeSignalMonitor.Dock();
            // if (APx.SignalMeters.IsDocked) APx.SignalMeters.Undock();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            //     createDataFolders();
            //     createLogFiles();       

            testResult0.Text = testResult1.Text = testResult2.Text = testResult3.Text = testResult4.Text = testResult5.Text = null;
            testResult6.Text = testResult7.Text = testResult8.Text = testResult9.Text = null;

            date_Stamp.Text = DateTime.Now.ToString("MM/dd/yyyy");
            time_Stamp.Text = DateTime.Now.ToString("hh:mm tt");

            pictureBox.Image = Image.FromFile(@"GENASYS Logo.bmp");

            setupInstruction.ForeColor = Color.Purple;

            setupRefData();
            setupProductModel(0, 0);
            displaySetupImages();
            DUT.portCState = 0x00;
            channelSetup(0x00, 0x00, DUT.portCState);
            resetSampleInfo();
            setupSerialNumberPerProduct();
            setupDriverSerialNumberPerProduct();
            recordSampleInfo(false);
            AUXCableOption.Visible = false;

            if (DUT.Debug == 0)
            {
                DebugMode.Text = null;
            }

            debugLabel1.Text = debugLabel2.Text = debugLabel3.Text = debugLabel4.Text = null;
            debugLabel5.Text = debugLabel6.Text = debugLabel7.Text = debugLabel8.Text = null;

        }

        void createDataFolders()
        {
            try
            {
                System.IO.Directory.CreateDirectory(Folder.networkDrive);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\nThis might be an issue with the network Connection\nPlease notify the Engineer in charge.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            System.IO.Directory.CreateDirectory(Folder.networkDrive + "Ref Files");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "100X");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "450XL");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "300X");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "1000");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "500X");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "1000X");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "2000X");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "1950XL");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "360X Manifolds");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "LRAD-RX");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "DS60");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "SSX");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "SS100");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "SS300");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "SS400");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "360Xm");
            System.IO.Directory.CreateDirectory(Folder.networkDrive + "360XL-MID");
        }

        void createLogFiles()
        {
            try
            {
                if (!File.Exists(Folder.networkDrive + "100X Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "100X Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Batt SN,Driver SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "100X Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "100X Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "100X Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Batt SN,Driver SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "100X Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "100X Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "100X Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Batt SN,Driver SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "100X Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "100X-NAVY-V1 Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "100X-NAVY-V1 Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Batt SN,Driver SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "100X-NAVY-V1 Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "100X-NAVY-V1 Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "100X-NAVY-V1 Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Batt SN,Driver SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "100X-NAVY-V1 Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "100X-NAVY-V1 Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "100X-NAVY-V1 Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Batt SN,Driver SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "100X-NAVY-V1 Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "100X-NAVY Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "100X-NAVY Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Elec SN,Mic SN,Batt SN,Driver SN,MP3-2 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,PWR-On/VB On,PWR-Off/VB Off,MP3-1 Sensitivity,Output Noise,IPOD Input,MaxSPL (dB),IPOD MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "100X-NAVY Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "100X-NAVY Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "100X-NAVY Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Elec SN,Mic SN,Batt SN,Driver SN,MP3-2 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,PWR-On/VB On,PWR-Off/VB Off,MP3-1 Sensitivity,Output Noise,IPOD Input,MaxSPL (dB),IPOD MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "100X-NAVY Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "100X-NAVY Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "100X-NAVY Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Elec SN,Mic SN,Batt SN,Driver SN,MP3-2 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,PWR-On/VB On,PWR-Off/VB Off,MP3-1 Sensitivity,Output Noise,IPOD Input,MaxSPL (dB),IPOD MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "100X-NAVY Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "300Xi Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "300Xi Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "300Xi Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "300Xi Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "300Xi Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "300Xi Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "300Xi Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "300Xi Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "300Xi Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "300XRA Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "300XRA Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "300XRA Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "300XRA Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "300XRA Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "300XRA Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "300XRA Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "300XRA Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "300XRA Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "300XRA-260W Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "300XRA-260W Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "300XRA-260W Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "300XRA-260W Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "300XRA-260W Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "300XRA-260W Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "300XRA-260W Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "300XRA-260W Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "300XRA-260W Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "450XL Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "450XL Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "450XL Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "450XL Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "450XL Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "450XL Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "450XL Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "450XL Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "450XL Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "450XL-RA Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "450XL-RA Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Driver 1 SN,Driver 2 SN,Remote Amp SN,Mic SN,Control Unit SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide-On,Narrow-Off,Narrow-On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "450XL-RA Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "450XL-RA Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "450XL-RA Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Driver 1 SN,Driver 2 SN,Remote Amp SN,Mic SN,Control Unit SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide-On,Narrow-Off,Narrow-On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "450XL-RA Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "450XL-RA Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "450XL-RA Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Driver 1 SN,Driver 2 SN,Remote Amp SN,Mic SN,Control Unit SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide-On,Narrow-Off,Narrow-On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "450XL-RA Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "500X Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "500X Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "500X Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "500X Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "500X Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "500X Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "500X Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "500X Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "500X Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "500X-RE Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "500X-RE Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "500X-RE Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "500X-RE Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "500X-RE Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "500X-RE Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "500X-RE Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "500X-RE Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "500X-RE Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000 Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000 Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("LRAD-1000-AHD-G,Operator,Date,Time,WO No,System P/F,1000,X-MP3-AL,1000-AMP,X-MIC-AL,1000-G-SYS,1000-AC-PWR,PHRASSL TR-2,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,Limited Key-SW,Self Test,AUX Cable,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000 Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000 Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000 Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("LRAD-1000-AHD-G,Operator,Date,Time,WO No,System P/F,1000,X-MP3-AL,1000-AMP,X-MIC-AL,1000-G-SYS,1000-AC-PWR,PHRASSL TR-2,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,Limited Key-SW,Self Test,AUX Cable,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000 Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000 Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000 Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("LRAD-1000-AHD-G,Operator,Date,Time,WO No,System P/F,1000,X-MP3-AL,1000-AMP,X-MIC-AL,1000-G-SYS,1000-AC-PWR,PHRASSL TR-2,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,Limited Key-SW,Self Test,AUX Cable,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000 Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000X Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000X Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Audio Output SW Off,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000X Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000X Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000X Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Audio Output SW Off,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000X Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000X Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000X Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Audio Output SW Off,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000X Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000X2U Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000X2U Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,2U Chassis,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mute Function,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000X2U Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000X2U Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000X2U Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,2U Chassis,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mute Function,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000X2U Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000X2U Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000X2U Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,2U Chassis,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mute Function,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000X2U Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000Xi Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000Xi Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Carbon-Fiber Hd SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000Xi Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000Xi Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000Xi Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Carbon-Fiber Hd SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000Xi Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000Xi Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000Xi Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Carbon-Fiber Hd SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000Xi Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000XVB Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000XVB Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Output SW-Max/VB-On,Output SW-Low/VB-Off,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000XVB Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000XVB Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000XVB Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Output SW-Max/VB-On,Output SW-Low/VB-Off,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000XVB Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1000XVB Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1000XVB Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,MP3 Player SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Output SW-Max/VB-On,Output SW-Low/VB-Off,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1000XVB Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "1950XL Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1950XL Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1950XL Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1950XL Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1950XL Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1950XL Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "1950XL Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "1950XL Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "1950XL Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "360X Manifold Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "360X Manifold Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "360X Manifold Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "360X Manifold Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "360X Manifold Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "360X Manifold Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "360X Manifold Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "360X Manifold Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "360X Manifold Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "60-DEG w AmpPack Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "60-DEG w AmpPack Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "60-DEG w AmpPack Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "60-DEG w AmpPack Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "60-DEG w AmpPack Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "60-DEG w AmpPack Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "60-DEG w AmpPack Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "60-DEG w AmpPack Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),NoiseLevel (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "60-DEG w AmpPack Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "60-DEG Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "60-DEG Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Transformer Box SN,Driver SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "60-DEG Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "60-DEG Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "60-DEG Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Transformer Box SN,Driver SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "60-DEG Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "60-DEG Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "60-DEG Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Transformer Box SN,Driver SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "60-DEG Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "60-DEG wo Amp Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "60-DEG wo Amp Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver SN,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "60-DEG wo Amp Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "60-DEG wo Amp Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "60-DEG wo Amp Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver SN,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "60-DEG wo Amp Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "60-DEG wo Amp Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "60-DEG wo Amp Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver SN,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "60-DEG wo Amp Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "SS100 Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SS100 Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver Panel SN,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SS100 Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SS100 Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SS100 Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver Panel SN,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SS100 Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SS100 Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SS100 Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver Panel SN,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SS100 Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "SS300 Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SS300 Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver Panel SN,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SS300 Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SS300 Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SS300 Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver Panel SN,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SS300 Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SS300 Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SS300 Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver Panel SN,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SS300 Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "SS400 Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SS400 Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver Panel SN,Sensitivity,Max SPL,Acelerometer,Output,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SS400 Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SS400 Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SS400 Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver Panel SN,Sensitivity,Max SPL,Acelerometer,Output,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SS400 Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SS400 Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SS400 Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver Panel SN,Sensitivity,Max SPL,Acelerometer,Output,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SS400 Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX60 Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX60 Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX60 Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX60 Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX60 Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX60 Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX60 Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX60 Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Transformer SN,HI-PWR,Max SPL,MID-PWR,LOW-PWR,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX60 Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX wo Trans Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX wo Trans Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX wo Trans Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX wo Trans Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX wo Trans Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX wo Trans Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX wo Trans Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX wo Trans Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX wo Trans Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX60 wo Trans Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX60 wo Trans Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX60 wo Trans Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX60 wo Trans Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX60 wo Trans Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX60 wo Trans Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "SSX60 wo Trans Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "SSX60 wo Trans Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "SSX60 wo Trans Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "LRAD-RX Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "LRAD-RX Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Firmware Ver,Elec SN,Head No 1 SN,Head No 2 SN,AMC No 1 SN,AMC No 2 SN,Light SN,Camera SN,48VDC Supply SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,Freq Sweep,MaxSPL,CoolingFansChk,Output Noise,Max SPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "LRAD-RX Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "LRAD-RX Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "LRAD-RX Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Firmware Ver,Elec SN,Head No 1 SN,Head No 2 SN,AMC No 1 SN,AMC No 2 SN,Light SN,Camera SN,48VDC Supply SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,Freq Sweep,MaxSPL,Cooling Fans,Max SPL (dB),Output Noise (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "LRAD-RX Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "LRAD-RX Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "LRAD-RX Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Firmware Ver,Elec SN,Head No 1 SN,Head No 2 SN,AMC No 1 SN,AMC No 2 SN,Light SN,Camera SN,48VDC Supply SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,Freq Sweep,MaxSPL,Cooling Fans,Max SPL (dB),Output Noise (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "LRAD-RX Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "2000X Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "2000X Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,Driver8 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "2000X Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "2000X Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "2000X Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,Driver8 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "2000X Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "2000X Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "2000X Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Control Unit SN,Elec SN,Mic SN,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,Driver5 SN,Driver6 SN,Driver7 SN,Driver8 SN,MP3 Sensitivity,Max SPL,Vol Function,Mic Sensitivity,Wide/On,Narrow/Off,Narrow/On,Output Noise,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "2000X Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "360XL-MID Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "360XL-MID Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver1 SN,Driver2 SN,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "360XL-MID Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "360XL-MID Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "360XL-MID Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver1 SN,Driver2 SN,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "360XL-MID Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "360XL-MID Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "360XL-MID Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver1 SN,Driver2 SN,Sensitivity,Max SPL,MaxSPL (dB),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "360XL-MID Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*************************************************************************************************************************************************************************/
            try
            {
                if (!File.Exists(Folder.networkDrive + "360Xm Test Log (Post-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "360Xm Test Log (Post-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,HI-PWR (TOP),Max SPL (TOP),MID-PWR (TOP),LOW-PWR (TOP),HI-PWR (BOT),Max SPL (BOT),MID-PWR (BOT),LOW-PWR (BOT),Max dBSPL (TOP),Max dBSPL (BOT),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "360Xm Test Log (Post-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "360Xm Test Log (Pre-Test).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "360Xm Test Log (Pre-Test).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,HI-PWR (TOP),Max SPL (TOP),MID-PWR (TOP),LOW-PWR (TOP),HI-PWR (BOT),Max SPL (BOT),MID-PWR (BOT),LOW-PWR (BOT),Max dBSPL (TOP),Max dBSPL (BOT),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "360Xm Test Log (Pre-Test).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
            try
            {
                if (!File.Exists(Folder.networkDrive + "360Xm Test Log (Repaired).csv"))
                {
                    StreamWriter LogFile = new StreamWriter(Folder.networkDrive + "360Xm Test Log (Repaired).csv", append: true);
                    LogFile.WriteLine("Unit SN,Operator,Date,Time,WO No,System P/F,Model,Driver1 SN,Driver2 SN,Driver3 SN,Driver4 SN,HI-PWR (TOP),Max SPL (TOP),MID-PWR (TOP),LOW-PWR (TOP),HI-PWR (BOT),Max SPL (BOT),MID-PWR (BOT),LOW-PWR (BOT),Max dBSPL (TOP),Max dBSPL (BOT),Software Ver,Remarks");
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + "360Xm Test Log (Repaired).csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //set version info
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = String.Format(this.Text, version.Major, version.Minor, version.Build, version.Revision);
        }

        private int getProductIndex()
        {
            if (productButton0.Checked)
            {
                return 0;
            }
            else if (productButton1.Checked)
            {
                return 1;
            }
            else if (productButton2.Checked)
            {
                return 2;
            }
            else if (productButton3.Checked)
            {
                return 3;
            }
            else if (productButton4.Checked)
            {
                return 4;
            }
            else if (productButton5.Checked)
            {
                return 5;
            }
            else if (productButton6.Checked)
            {
                return 6;
            }
            else if (productButton7.Checked)
            {
                return 7;
            }
            else if (productButton8.Checked)
            {
                return 8;
            }
            else if (productButton9.Checked)
            {
                return 9;
            }
            else if (productButton10.Checked)
            {
                return 10;
            }
            else if (productButton11.Checked)
            {
                return 11;
            }
            else if (productButton12.Checked)
            {
                return 12;
            }
            else if (productButton13.Checked)
            {
                return 13;
            }
            else if (productButton14.Checked)
            {
                return 14;
            }
            else if (productButton15.Checked)
            {
                return 15;
            }
            else if (productButton16.Checked)
            {
                return 16;
            }
            else if (productButton17.Checked)
            {
                return 17;
            }
            else if (productButton18.Checked)
            {
                return 18;
            }
            else if (productButton19.Checked)
            {
                return 19;
            }
            else if (productButton20.Checked)
            {
                return 20;
            }
            else if (productButton21.Checked)
            {
                return 21;
            }
            else if (productButton22.Checked)
            {
                return 22;
            }
            else
            {
                return -1;
            }
        }

        private int getModelIndex()
        {
            if (modelButton0.Checked)
            {
                return 0;
            }
            else if (modelButton1.Checked)
            {
                return 1;
            }
            else if (modelButton2.Checked)
            {
                return 2;
            }
            else if (modelButton3.Checked)
            {
                return 3;
            }
            else if (modelButton4.Checked)
            {
                return 4;
            }
            else if (modelButton5.Checked)
            {
                return 5;
            }
            else
            {
                return -1;
            }
        }

        public void channelSetup(byte portA, byte portB, byte portC)
        {
            /*   AIOWDM.RelOutPortB(CardNum, Offset, portA);
               AIOWDM.RelOutPortB(CardNum, Offset + 1, portB);
               AIOWDM.RelOutPortB(CardNum, Offset + 2, portC);
               Thread.Sleep(10);
               getChannelStatus();  */ //TempDebug
            Thread.Sleep(10);
        }

        public int getChannelStatus()
        {
            UInt16 PortA, PortB, PortC = 0;
            int testCh = 0;

            /*   PortA = AIOWDM.RelInPortB(CardNum, Offset);
               PortB = AIOWDM.RelInPortB(CardNum, Offset + 1);
               PortC = AIOWDM.RelInPortB(CardNum, Offset + 2);   */ //TempDebug

            /*    if (DUT.Debug == 1)
                {
                    debugLabel3.Text = "Black Box Status: PortA = " + Convert.ToString(PortA) + ", PortB = " + Convert.ToString(PortB) + ", PortC = " + Convert.ToString(PortC);
                    debugLabel3.Update();
                }   */ //TempDebug

            return testCh;
        }

        void setupProductModel(int productIndex, int modelIndex)
        {
            modelButton0.Visible = modelButton1.Visible = modelButton2.Visible = modelButton3.Visible = modelButton4.Visible = modelButton5.Visible = false;
            modelButton0.Enabled = modelButton1.Enabled = modelButton2.Enabled = modelButton3.Enabled = modelButton4.Enabled = modelButton5.Enabled = true;

            DUT.ExtAmplifier = 0;

            if (productIndex == 0)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton2.Visible = true;
                modelButton0.Text = "100X";
                modelButton1.Text = "100X-NAVY V01";
                modelButton2.Text = "100X-NAVY";

                if (modelIndex > 2)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "100X";
                }
                else if (modelIndex == 0) DUT.productModel = "100X";
                else if (modelIndex == 1) DUT.productModel = "100X-NAVY-V1";
                else if (modelIndex == 2) DUT.productModel = "100X-NAVY";
            }
            else if (productIndex == 1)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton2.Visible = true;
                modelButton0.Text = "300Xi";
                modelButton1.Text = "300X-RA";
                modelButton2.Text = "300XRA-260W";

                if (modelIndex > 2)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "300X";
                }
                else if (modelIndex == 0) DUT.productModel = "300Xi";
                else if (modelIndex == 1) DUT.productModel = "300XRA";
                else if (modelIndex == 2) DUT.productModel = "300XRA-260W";
            }
            else if (productIndex == 2)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton2.Visible = true;
                modelButton0.Text = "450XL";
                modelButton1.Text = "450XL Extended Test";
                modelButton2.Text = "450XL-RA";

                if (modelIndex > 2)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "450XL";
                }
                else if (modelIndex == 0) DUT.productModel = "450XL";
                else if (modelIndex == 1) DUT.productModel = "450XL Extended Test";
                else if (modelIndex == 2) DUT.productModel = "450XL-RA";
            }
            else if (productIndex == 3)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton0.Text = "500X";
                modelButton1.Text = "500XRE";

                if (modelIndex > 1)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "500X";
                }
                else if (modelIndex == 0) DUT.productModel = "500X";
                else if (modelIndex == 1) DUT.productModel = "500X-RE";
            }
            else if (productIndex == 4)
            {
                DUT.productModel = "1000";
            }
            else if (productIndex == 5)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton2.Visible = true;
                modelButton0.Text = "20' Audio Cable";
                modelButton1.Text = "30' Audio Cable";
                modelButton2.Text = "100' Audio Cable";

                if (modelIndex > 2)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "1000X-20FT";
                }
                else if (modelIndex == 0) DUT.productModel = "1000X-20FT";
                else if (modelIndex == 1) DUT.productModel = "1000X-30FT";
                else if (modelIndex == 2) DUT.productModel = "1000X-100FT";
            }
            else if (productIndex == 6)
            {
                DUT.productModel = "1000X2U";
            }
            else if (productIndex == 7)
            {
                DUT.productModel = "1000Xi";
            }
            else if (productIndex == 8)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton2.Visible = true;
                modelButton0.Text = "20' Audio Cable";
                modelButton1.Text = "30' Audio Cable";
                modelButton2.Text = "100' Audio Cable";

                if (modelIndex > 2)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "1000XVB-20FT";
                }
                else if (modelIndex == 0) DUT.productModel = "1000XVB-20FT";
                else if (modelIndex == 1) DUT.productModel = "1000XVB-30FT";
                else if (modelIndex == 2) DUT.productModel = "1000XVB-100FT";
            }
            else if (productIndex == 9)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton2.Visible = true;
                modelButton0.Text = "35' Audio Cable";
                modelButton1.Text = "66' Audio Cable";
                modelButton2.Text = "100' Audio Cable";

                if (modelIndex > 2)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "1950XL-35FT";
                }
                else if (modelIndex == 0) DUT.productModel = "1950XL-35FT";
                else if (modelIndex == 1) DUT.productModel = "1950XL-66FT";
                else if (modelIndex == 2) DUT.productModel = "1950XL-100FT";
            }
            else if (productIndex == 10)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton0.Text = "Standard Drivers";
                modelButton1.Text = "Enhanced Drivers";

                if (modelIndex > 1)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "360X Manifold";
                }
                else if (modelIndex == 0) DUT.productModel = "360X Manifold";
                else if (modelIndex == 1) DUT.productModel = "360XL Manifold";
            }
            else if (productIndex == 11)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton0.Text = "Standard Drivers";
                modelButton1.Text = "Enhanced Drivers";

                if (modelIndex > 1)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "DS60-X";
                }
                else if (modelIndex == 0) DUT.productModel = "DS60-X w AmpPack";
                else if (modelIndex == 1) DUT.productModel = "DS60-XL w AmpPack";
            }
            else if (productIndex == 12)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton2.Visible = true;
                modelButton3.Visible = true;
                modelButton4.Visible = true;
                modelButton0.Text = "DS60-70V-60W";
                modelButton1.Text = "DS60-100V-60W";
                modelButton2.Text = "DS60-25V-80W";
                modelButton3.Text = "DS60-70V-80W";
                modelButton4.Text = "DS60-100V-80W";

                DUT.ExtAmplifier = 1;

                if (modelIndex > 4)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "DS60-70V-60W";
                }
                else if (modelIndex == 0) DUT.productModel = "DS60-70V-60W";
                else if (modelIndex == 1) DUT.productModel = "DS60-100V-60W";
                else if (modelIndex == 2) DUT.productModel = "DS60-25V-80W";
                else if (modelIndex == 3) DUT.productModel = "DS60-70V-80W";
                else if (modelIndex == 4) DUT.productModel = "DS60-100V-80W";
            }
            else if (productIndex == 13)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton0.Text = "DS60-70V-160W";
                modelButton1.Text = "DS60-100V-160W";

                DUT.ExtAmplifier = 1;

                if (modelIndex > 1)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "DS60-70V-160W";
                }
                else if (modelIndex == 0) DUT.productModel = "DS60-70V-160W";
                else if (modelIndex == 1) DUT.productModel = "DS60-100V-160W";
            }
            else if (productIndex == 14)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton0.Text = "DS60-X";
                modelButton1.Text = "DS60-XL";

                DUT.ExtAmplifier = 1;

                if (modelIndex > 1)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "DS60-X";
                }
                else if (modelIndex == 0) DUT.productModel = "DS60-X";
                else if (modelIndex == 1) DUT.productModel = "DS60-XL";
            }
            else if (productIndex == 15)
            {
                modelButton0.Visible = modelButton1.Visible = modelButton2.Visible = true;

                //       if (DUT.EngTestRoomSetup == 1) { modelButton0.Enabled = modelButton2.Enabled = false; }
                //     else { modelButton1.Enabled = false; }

                modelButton0.Text = "SS100";
                modelButton1.Text = "SS300";
                modelButton2.Text = "SS400";

                DUT.ExtAmplifier = 1;

                if (modelIndex > 2)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "SS100";
                }
                else if (modelIndex == 0) DUT.productModel = "SS100";
                else if (modelIndex == 1) DUT.productModel = "SS300";
                else if (modelIndex == 2) DUT.productModel = "SS400";
            }
            else if (productIndex == 16)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton0.Text = "SS-X120";
                modelButton1.Text = "SS-X60";

                DUT.ExtAmplifier = 1;

                if (modelIndex > 1)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "SSX";
                }
                else if (modelIndex == 0) DUT.productModel = "SSX";
                else if (modelIndex == 1) DUT.productModel = "SSX60";
            }
            else if (productIndex == 17)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton0.Text = "SS-X120";
                modelButton1.Text = "SS-X60";

                DUT.ExtAmplifier = 1;

                if (modelIndex > 1)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "SSX wo Trans";
                }
                else if (modelIndex == 0) DUT.productModel = "SSX wo Trans";
                else if (modelIndex == 1) DUT.productModel = "SSX60 wo Trans";
            }
            else if (productIndex == 18)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton2.Visible = true;
                modelButton3.Visible = true;
                modelButton0.Enabled = true;
                modelButton1.Enabled = true;
                modelButton2.Enabled = true;
                modelButton3.Enabled = true;
                modelButton0.Text = "500RX";
                modelButton1.Text = "950RXL";
                modelButton2.Text = "1000RX";
                modelButton3.Text = "950NXT";

                if (modelIndex > 3)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "500RX";
                }
                else if (modelIndex == 0) DUT.productModel = "500RX";
                else if (modelIndex == 1) DUT.productModel = "950RXL";
                else if (modelIndex == 2) DUT.productModel = "1000RX";
                else if (modelIndex == 3) DUT.productModel = "950NXT";
            }
            else if (productIndex == 19)
            {
                DUT.productModel = "2000X";
            }
            else if (productIndex == 20)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton2.Visible = true;
                modelButton3.Visible = true;
                modelButton4.Visible = true;
                modelButton5.Visible = true;
                modelButton0.Text = "1-ST 25V-60W";
                modelButton1.Text = "1-ST 100V-60W";
                modelButton2.Text = "2-ST 25V-60W";
                modelButton3.Text = "2-ST 100V-120W";
                modelButton4.Text = "2-ST 70V-60W";
                modelButton5.Text = "4-ST 100V-240W";

                DUT.ExtAmplifier = 1;

                if (modelIndex > 5)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "360Xm 1-ST 25V-60W";
                }
                else if (modelIndex == 0) DUT.productModel = "360Xm 1-ST 25V-60W";
                else if (modelIndex == 1) DUT.productModel = "360Xm 1-ST 100V-60W";
                else if (modelIndex == 2) DUT.productModel = "360Xm 2-ST 25V-60W";
                else if (modelIndex == 3) DUT.productModel = "360Xm 2-ST 100V-120W";
                else if (modelIndex == 4) DUT.productModel = "360Xm 2-ST 70V-60W";
                else if (modelIndex == 5) DUT.productModel = "360Xm 4-ST 100V-240W";
            }
            else if (productIndex == 21)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton0.Text = "1-ST wo Trans";
                modelButton1.Text = "2-ST wo Trans";

                DUT.ExtAmplifier = 1;

                if (modelIndex > 1)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "360Xm 1-ST wo Trans";
                }
                else if (modelIndex == 0) DUT.productModel = "360Xm 1-ST wo Trans";
                else if (modelIndex == 1) DUT.productModel = "360Xm 2-ST wo Trans";
            }
            else if (productIndex == 22)
            {
                modelButton0.Visible = true;
                modelButton1.Visible = true;
                modelButton0.Text = "360XL-MID 1-ST";
                modelButton1.Text = "360XL-MID 2-ST";

                DUT.ExtAmplifier = 1;

                if (modelIndex > 1)
                {
                    modelButton0.Checked = true;
                    DUT.productModel = "360X-MID 1-ST";
                }
                else if (modelIndex == 0) DUT.productModel = "360XL-MID 1-ST";
                else if (modelIndex == 1) DUT.productModel = "360XL-MID 2-ST";
            }

            if (DUT.Debug == 1) { }

        }

        private void SN_Changed(object sender, EventArgs e)
        {
            //*** Save the last data  //here sp3
            SaveData(DUT.FileNameExt);

            //troubleShootingCheckBox.Checked = false;
            remarkBox.Clear();
            DUT.resetVariablesToDefault();
            setupTestSelection();
            resetreportFormat();
        }

        private void productModelChanged(object sender, EventArgs e)
        {
            SaveData(DUT.FileNameExt);
            shutdownTimer.Stop();

            if (DUT.autoShutdown == 1)
            {
                shutdownTimer.Start();
            }

            DUT.resetVariablesToDefault();
            clearAllRemarks();
            troubleShootingCheckBox.Checked = false;
            AUXCableOption.Visible = false;
            AUXCableOption.Checked = false;
            postTest.Checked = true;
            groupBox8.Update();

            setupProductModel(getProductIndex(), getModelIndex()); //step1  
            setupTestSelection();
            displaySetupImages();
            resetSampleInfo();
            setupSerialNumberPerProduct();
            setupDriverSerialNumberPerProduct();
            setupRefData();
            recordSampleInfo(false);

            if (DUT.Debug == 0)
            {
                DebugMode.Text = null;
            }
            else
            {
                DebugMode.Text = "Debug Mode Active";
            }

            if ((DUT.productModel == "SS300") || (productButton20.Checked)) //Using test cart /w trans selectable outputs
            {
                DUT.portCState = (byte)(DUT.portCState | 8); //Disconnect 5-Ohm load
                channelSetup(0x00, 0x00, DUT.portCState);
            }
            else
            {
                Set_LRADX_MP3_LowerMic();
            }

            DisplayDebugInfo();

        } //end of productModelChanged(object sender, EventArgs e)

        void clearAllRemarks()
        {
            remarkBox.Clear();
            remarkBox1.Clear();
            remarkBox2.Clear();
            remarkBox3.Clear();
            remarkBox4.Clear();
            remarkBox5.Clear();
            remarkBox6.Clear();
            remarkBox7.Clear();
            remarkBox8.Clear();
            remarkBox9.Clear();
            remarkBox10.Clear();
            remarkBox11.Clear();
            remarkBox12.Clear();
            remarkBox13.Clear();
            remarkBox14.Clear();
            remarkBox15.Clear();
            remarkBox16.Clear();
            remarkBox17.Clear();
            remarkBox18.Clear();
        }

        void resetSampleInfo()
        {
            WO_No.Text = null;
            SN_Box1.Text = SN_Box2.Text = SN_Box3.Text = SN_Box4.Text = SN_Box5.Text = null;
            SN_Box6.Text = SN_Box7.Text = null;

            driverSN_Box1.Text = driverSN_Box2.Text = driverSN_Box3.Text = driverSN_Box4.Text = null;
            driverSN_Box5.Text = driverSN_Box6.Text = driverSN_Box7.Text = driverSN_Box8.Text = null;

            resetreportFormat();

        }

        void resetreportFormat()
        {
            reportAudioResult0.Text = null;
            reportAudioResult1.Text = null;
            reportAudioResult2.Text = null;
            reportAudioResult3.Text = null;
            reportAudioResult4.Text = null;
            reportAudioResult5.Text = null;
            reportAudioResult6.Text = null;

            reportSerialNo1.Text = reportSerialNo2.Text = reportSerialNo3.Text = reportSerialNo4.Text = null;
            reportSerialNo5.Text = reportSerialNo6.Text = reportSerialNo7.Text = null;

            reportPassFail0.Text = reportPassFail1.Text = reportPassFail2.Text = reportPassFail3.Text = null;
            reportPassFail4.Text = reportPassFail5.Text = reportPassFail6.Text = null;

            /*******************************************************************************************************************************************************************************/
            /* 3rd sesion: Test Results
            /*******************************************************************************************************************************************************************************/
            if ((productButton12.Checked) || (productButton13.Checked) || (productButton16.Checked) || (DUT.productModel == "SS100") || (DUT.productModel == "SS300") ||
              (DUT.productModel == "360Xm 1-ST 100V-60W") || (DUT.productModel == "360Xm 1-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 100V-120W") || (DUT.productModel == "360Xm 4-ST w Trans")) //DS60s w Trans/SSX/SS100/SS300
            {
                reportAudioResult0.Text = "  HIGH Power Output";
                reportAudioResult1.Text = "  MID Power Output";
                reportAudioResult2.Text = "  LOW Power Output";
                reportAudioResult3.Text = "  Maximum SPL";

                reportPassFail0.Text = reportPassFail1.Text = reportPassFail2.Text = reportPassFail3.Text = "N/A";
            }
            else if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W")) //360Xm w 2 separated stacks
            {
                reportAudioResult0.Text = "  HIGH Power Output (TOP)";
                reportAudioResult1.Text = "  MID Power Output (TOP)";
                reportAudioResult2.Text = "  LOW Power Output (TOP)";
                reportAudioResult3.Text = "  HIGH Power Output (BOT)";
                reportAudioResult4.Text = "  MID Power Output (BOT)";
                reportAudioResult5.Text = "  LOW Power Output (BOT)";

                reportPassFail0.Text = reportPassFail1.Text = reportPassFail2.Text = "N/T";
                reportPassFail3.Text = reportPassFail4.Text = reportPassFail5.Text = "N/T";
            }
            else if (DUT.productModel == "360Xm 4-ST 100V-240W")
            {
                reportAudioResult0.Text = "  Sensitivity";
                reportAudioResult1.Text = "  Maximum SPL";
            }
            else if (productButton18.Checked)
            {
                reportAudioResult0.Text = "  Frequency Sweep";
                reportAudioResult1.Text = "  Maximum SPL";
                reportAudioResult3.Text = "  Maximum Noise Level Output";
                reportPassFail0.Text = reportPassFail1.Text = reportPassFail3.Text = "N/T";

                if (DUT.productModel != "950NXT")
                {
                    reportAudioResult2.Text = "  Cooling Fans Function";
                    reportPassFail2.Text = reportPassFail3.Text = "N/T";
                }

            }
            else if (DUT.productModel == "SS400")
            {
                reportAudioResult0.Text = "  Output Sensitivity/Max SPL";
                reportAudioResult1.Text = "  Accelerometer Output";

                reportPassFail0.Text = reportPassFail1.Text = "N/T";
            }
            else if ((productButton10.Checked) || (productButton14.Checked) || (productButton17.Checked)) //Manifolds/SSX wo Trans/DS60 Horns only
            {
                reportAudioResult0.Text = "  Output Sensitivity";
                reportAudioResult1.Text = "  Maximum SPL";

                reportPassFail0.Text = reportPassFail1.Text = "N/T";
            }
            else if ((DUT.productModel == "100X") || (DUT.productModel == "100X-NAVY-V1") || (DUT.productModel == "1000X2U"))
            {
                reportAudioResult0.Text = "  MP3 Input Sensitivity / Maximum SPL";
                reportAudioResult1.Text = "  Volume Control Function";
                reportAudioResult3.Text = "  Maximum Noise level output";

                if (DUT.productModel == "1000X2U") reportAudioResult2.Text = "  MUTE Function";
                else reportAudioResult2.Text = "  MIC Input Sensitivity";

                reportPassFail0.Text = reportPassFail1.Text = reportPassFail2.Text = reportPassFail3.Text = "N/T";
            }
            else if (DUT.productModel == "100X-NAVY")
            {
                reportAudioResult0.Text = "  MP3-2 Input Sensitivity / Maximum SPL";
                reportAudioResult1.Text = "  Volume Control Function";
                reportAudioResult2.Text = "  MIC Input Sensitivity";
                reportAudioResult3.Text = "  Audio Effect Controls";
                reportAudioResult4.Text = "  MP3-1 Input Sensitivity";
                reportAudioResult5.Text = "  Maximum Noise level output";
                reportAudioResult6.Text = "  IPOD Input Sensitivity / Maximum SPL";

                reportPassFail0.Text = reportPassFail1.Text = reportPassFail2.Text = reportPassFail3.Text = reportPassFail4.Text = reportPassFail5.Text = reportPassFail6.Text = "N/T";
            }
            else if (DUT.productModel == "1000X2U")
            {
                reportAudioResult0.Text = "  MP3 Input Sensitivity / Maximum SPL";
                reportAudioResult1.Text = "  Volume Control Function";
                reportAudioResult2.Text = "  MUTE Function";
                reportAudioResult3.Text = "  Maximum Noise level output";

                reportPassFail0.Text = reportPassFail1.Text = reportPassFail2.Text = reportPassFail3.Text = "N/T";
            }
            else if (productButton22.Checked) //360X-MIDs
            {
                reportAudioResult0.Text = "  Sensitivity";
                reportAudioResult1.Text = "   Maximum SPL";
                //   reportAudioResult2.Text = ""; reportAudioResult3.Text = "";

                reportPassFail0.Text = reportPassFail1.Text = "N/T";
            }
            else
            {
                if (productButton4.Checked) //1000V
                {
                    reportAudioResult5.Text = "  Limited Switch & AUX Cable";
                    reportAudioResult6.Text = "  Self Test Function";
                    reportPassFail5.Text = reportPassFail6.Text = reportPassFail2.Text = "N/T";
                }
                reportAudioResult0.Text = "  MP3 Input Sensitivity / Maximum SPL";
                reportAudioResult1.Text = "  Volume Control Function";
                reportAudioResult2.Text = "  MIC Input Sensitivity";
                reportAudioResult3.Text = "  Audio Effect Controls";
                reportAudioResult4.Text = "  Maximum Noise level output";

                reportPassFail0.Text = reportPassFail1.Text = reportPassFail2.Text = "N/T";
                reportPassFail3.Text = reportPassFail4.Text = "N/T";
            }

            reportAccessoryNo1.Text = reportAccessoryNo2.Text = reportAccessoryNo3.Text = reportAccessoryNo4.Text = null;
        }

        void setupTestSelection()
        {
            testSelButton0.Visible = false;
            testSelButton1.Visible = false;
            testSelButton2.Visible = false;
            testSelButton3.Visible = false;
            testSelButton4.Visible = false;
            testSelButton5.Visible = false;
            testSelButton6.Visible = false;
            testSelButton7.Visible = false;
            testSelButton8.Visible = false;
            testSelButton9.Visible = false;
            testSelButton0.Checked = true; // troubleShootingCheckBox.Checked = false;

            testResult0.Text = testResult1.Text = testResult2.Text = testResult3.Text = testResult4.Text = testResult5.Text = testResult6.Text = null;
            testResult7.Text = testResult8.Text = testResult9.Text = null;

            setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction0");
            if (setupInstruction.Text == "")
            {
                setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction0");
            }

            if ((DUT.productModel == "100X") || (DUT.productModel == "100X-NAVY-V1"))
            {
                testSelButton0.Visible = true;
                testSelButton1.Visible = true;
                testSelButton2.Visible = true;
                testSelButton6.Visible = true;
                testSelButton0.Text = "MP3 Input Sensitivity / Max. SPL -------------------------------------------------------------------------------------------";
                testSelButton1.Text = "Volume Control Knob Function -----------------------------------------------------------------------------------------------";
                testSelButton2.Text = "Microphone Input Sensitivity -----------------------------------------------------------------------------------------------";
                testSelButton6.Text = "System Background Noise Output -----------------------------------------------------------------------------------------------";

                testResult0.ForeColor = Color.LightGray;
                testResult0.ForeColor = Color.LightGray;
                testResult0.Text = "NT";
                testResult1.ForeColor = Color.LightGray;
                testResult1.ForeColor = Color.LightGray;
                testResult1.Text = "NT";
                testResult2.ForeColor = Color.LightGray;
                testResult2.ForeColor = Color.LightGray;
                testResult2.Text = "NT";
                testResult6.ForeColor = Color.LightGray;
                testResult8.ForeColor = Color.LightGray;
                testResult6.Text = "NT";

            }
            else if (DUT.productModel == "100X-NAVY")
            {
                testSelButton0.Visible = true;
                testSelButton1.Visible = true;
                testSelButton2.Visible = true;
                testSelButton3.Visible = true;
                testSelButton4.Visible = true;
                testSelButton5.Visible = true;
                testSelButton6.Visible = true;
                testSelButton7.Visible = true;
                testSelButton0.Text = "MP3 Input2 Sensitivity / Max. SPL -------------------------------------------------------------------------------------------";
                testSelButton1.Text = "Volume Control Knob Function -----------------------------------------------------------------------------------------------";
                testSelButton2.Text = "Microphone Input Sensitivity -----------------------------------------------------------------------------------------------";
                testSelButton3.Text = "Audio Effect with PWR SW ON, VB ON ----------------------------------------------------------------------------------------------------";
                testSelButton4.Text = "Audio Effect with PWR SW OFF, VB OFF ---------------------------------------------------------------------------------------------------";
                testSelButton5.Text = "MP3 Input1 Sensitivity -----------------------------------------------------------------------------------------------------";
                testSelButton6.Text = "System Background Noise Output -----------------------------------------------------------------------------------------------";
                testSelButton7.Text = "IPOD Input Sensitivity / Max. SPL -----------------------------------------------------------------------------------------------";

                testResult0.ForeColor = Color.LightGray;
                testResult0.ForeColor = Color.LightGray;
                testResult0.Text = "NT";
                testResult1.ForeColor = Color.LightGray;
                testResult1.ForeColor = Color.LightGray;
                testResult1.Text = "NT";
                testResult2.ForeColor = Color.LightGray;
                testResult2.ForeColor = Color.LightGray;
                testResult2.Text = "NT";
                testResult3.ForeColor = Color.LightGray;
                testResult3.ForeColor = Color.LightGray;
                testResult3.Text = "NT";
                testResult4.ForeColor = Color.LightGray;
                testResult4.ForeColor = Color.LightGray;
                testResult4.Text = "NT";
                testResult5.ForeColor = Color.LightGray;
                testResult5.ForeColor = Color.LightGray;
                testResult5.Text = "NT";
                testResult6.ForeColor = Color.LightGray;
                testResult6.ForeColor = Color.LightGray;
                testResult6.Text = "NT";
                testResult7.ForeColor = Color.LightGray;
                testResult7.ForeColor = Color.LightGray;
                testResult7.Text = "NT";
            }
            else if ((productButton1.Checked) || (productButton2.Checked) || (productButton3.Checked) || (productButton7.Checked) || (productButton9.Checked) || (productButton11.Checked))
            {
                testSelButton0.Visible = true;
                testSelButton1.Visible = true;
                testSelButton2.Visible = true;
                testSelButton3.Visible = true;
                testSelButton4.Visible = true;
                testSelButton5.Visible = true;
                testSelButton6.Visible = true;
                testSelButton0.Text = "MP3 Input Sensitivity / Max. SPL -------------------------------------------------------------------------------------------";
                testSelButton1.Text = "Volume Control Knob Function -----------------------------------------------------------------------------------------------";
                testSelButton2.Text = "Microphone Input Sensitivity -----------------------------------------------------------------------------------------------";
                testSelButton3.Text = "Audio Effect /w Sound=Wide, VB=ON -------------------------------------------------------------------------------------------";
                testSelButton4.Text = "Audio Effect /w Sound=Narrow, VB=OFF ----------------------------------------------------------------------------------------";
                testSelButton5.Text = "Audio Effect /w Sound=Narrow, VB=ON -----------------------------------------------------------------------------------------";
                testSelButton6.Text = "System Background Noise Output ----------------------------------------------------------------------------------------------";

                testResult0.ForeColor = Color.LightGray;
                testResult0.ForeColor = Color.LightGray;
                testResult0.Text = "NT";
                testResult1.ForeColor = Color.LightGray;
                testResult1.ForeColor = Color.LightGray;
                testResult1.Text = "NT";
                testResult2.ForeColor = Color.LightGray;
                testResult2.ForeColor = Color.LightGray;
                testResult2.Text = "NT";
                testResult3.ForeColor = Color.LightGray;
                testResult3.ForeColor = Color.LightGray;
                testResult3.Text = "NT";
                testResult4.ForeColor = Color.LightGray;
                testResult4.ForeColor = Color.LightGray;
                testResult4.Text = "NT";
                testResult5.ForeColor = Color.LightGray;
                testResult5.ForeColor = Color.LightGray;
                testResult5.Text = "NT";
                testResult6.ForeColor = Color.LightGray;
                testResult6.ForeColor = Color.LightGray;
                testResult6.Text = "NT";
            }
            else if ((productButton5.Checked) || (productButton8.Checked)) //1000X/1000XVB
            { //sp
                testSelButton0.Visible = true;
                testSelButton1.Visible = true;
                testSelButton2.Visible = true;
                testSelButton3.Visible = true;
                testSelButton6.Visible = true;
                testSelButton0.Text = "MP3 Input Sensitivity / Max. SPL -------------------------------------------------------------------------------------------";
                testSelButton1.Text = "Volume Control Knob Function -----------------------------------------------------------------------------------------------";
                testSelButton2.Text = "Microphone Input Sensitivity ------------------------------------------------------------------------------------------------";
                testSelButton6.Text = "System Background Noise Output ----------------------------------------------------------------------------------------------";

                testResult0.ForeColor = Color.LightGray;
                testResult0.ForeColor = Color.LightGray;
                testResult0.Text = "NT";
                testResult1.ForeColor = Color.LightGray;
                testResult1.ForeColor = Color.LightGray;
                testResult1.Text = "NT";
                testResult2.ForeColor = Color.LightGray;
                testResult2.ForeColor = Color.LightGray;
                testResult2.Text = "NT";
                testResult6.ForeColor = Color.LightGray;
                testResult8.ForeColor = Color.LightGray;
                testResult6.Text = "NT";

                if (productButton5.Checked)
                {
                    testSelButton3.Text = "Output Sensitivity /w Audio Output Switch = LOW -------------------------------------------------------------------------------------";
                    testResult3.ForeColor = Color.LightGray;
                    testResult3.ForeColor = Color.LightGray;
                    testResult3.Text = "NT";
                }
                else
                {
                    testSelButton4.Visible = true;
                    testSelButton3.Text = "Output Sensitivity /w Audio Output Switch = MAX., VB = ON -------------------------------------------------------------------------------------";
                    testSelButton4.Text = "Output Sensitivity /w Audio Output Switch = LOW, VB = OFF -------------------------------------------------------------------------------------";
                    testResult3.ForeColor = Color.LightGray;
                    testResult3.ForeColor = Color.LightGray;
                    testResult3.Text = "NT";
                    testResult4.ForeColor = Color.LightGray;
                    testResult4.ForeColor = Color.LightGray;
                    testResult4.Text = "NT";
                }
            }
            else if (DUT.productModel == "1000X2U")
            {
                testSelButton0.Visible = true;
                testSelButton1.Visible = true;
                testSelButton2.Visible = true;
                testSelButton6.Visible = true;
                testSelButton0.Text = "Input Sensitivity / Max. SPL -------------------------------------------------------------------------------------------";
                testSelButton1.Text = "Volume Control Knob Function -----------------------------------------------------------------------------------------------";
                testSelButton2.Text = "MUTE Function ------------------------------------------------------------------------------------------------";
                testSelButton6.Text = "System Background Noise Output ----------------------------------------------------------------------------------------------";

                testResult0.ForeColor = Color.LightGray;
                testResult0.ForeColor = Color.LightGray;
                testResult0.Text = "NT";
                testResult1.ForeColor = Color.LightGray;
                testResult1.ForeColor = Color.LightGray;
                testResult1.Text = "NT";
                testResult2.ForeColor = Color.LightGray;
                testResult2.ForeColor = Color.LightGray;
                testResult2.Text = "NT";
                testResult6.ForeColor = Color.LightGray;
                testResult8.ForeColor = Color.LightGray;
                testResult6.Text = "NT";
            }
            else if ((productButton12.Checked) || (productButton13.Checked) || (productButton16.Checked) || (DUT.productModel == "SS100") ||
              (productButton20.Checked) || (DUT.productModel == "SS300")) //DS60-X, DS60-XL, SS-X60 & SS-X120, SS100/300
            {
                testResult0.ForeColor = Color.LightGray;
                testResult0.ForeColor = Color.LightGray;
                testResult0.Text = "NT";

                if ((modelButton2.Checked) || (modelButton4.Checked))
                {
                    testSelButton3.Visible = true;
                    testSelButton4.Visible = true;
                    testSelButton5.Visible = true;
                    testSelButton0.Visible = true;
                    testSelButton1.Visible = true;
                    testSelButton2.Visible = true;
                    testSelButton0.Text = "HIGH Power Sensitivity / Max. SPL (TOP) -------------------------------------------------------------------------------------------";
                    testSelButton1.Text = "MID Power Sensitivity (TOP) -----------------------------------------------------------------------------------------------";
                    testSelButton2.Text = "LOW Power Sensitivity  (TOP) ------------------------------------------------------------------------------------------------";
                    testSelButton3.Text = "HIGH Power Sensitivity / Max. SPL (BOT) -------------------------------------------------------------------------------------------";
                    testSelButton4.Text = "MID Power Sensitivity (BOT) -----------------------------------------------------------------------------------------------";
                    testSelButton5.Text = "LOW Power Sensitivity (BOT) ------------------------------------------------------------------------------------------------";

                    testResult1.ForeColor = Color.LightGray;
                    testResult1.ForeColor = Color.LightGray;
                    testResult1.Text = "NT";
                    testResult2.ForeColor = Color.LightGray;
                    testResult2.ForeColor = Color.LightGray;
                    testResult2.Text = "NT";
                    testResult3.ForeColor = Color.LightGray;
                    testResult3.ForeColor = Color.LightGray;
                    testResult3.Text = "NT";
                    testResult4.ForeColor = Color.LightGray;
                    testResult4.ForeColor = Color.LightGray;
                    testResult4.Text = "NT";
                    testResult5.ForeColor = Color.LightGray;
                    testResult5.ForeColor = Color.LightGray;
                    testResult5.Text = "NT";
                }
                else if (DUT.productModel == "360Xm 4-ST 100V-240W")
                {
                    testSelButton0.Visible = true;
                    testSelButton0.Text = "Sensitivity / Max. SPL -------------------------------------------------------------------------------------------";

                }
                else
                {
                    testSelButton0.Visible = true;
                    testSelButton1.Visible = true;
                    testSelButton2.Visible = true;
                    testSelButton0.Text = "HIGH Power Sensitivity / Max. SPL -------------------------------------------------------------------------------------------";
                    testSelButton1.Text = "MID Power Sensitivity -----------------------------------------------------------------------------------------------";
                    testSelButton2.Text = "LOW Power Sensitivity  ------------------------------------------------------------------------------------------------";

                    testResult1.ForeColor = Color.LightGray;
                    testResult1.ForeColor = Color.LightGray;
                    testResult1.Text = "NT";
                    testResult2.ForeColor = Color.LightGray;
                    testResult2.ForeColor = Color.LightGray;
                    testResult2.Text = "NT";
                }

            }
            else if (DUT.productModel == "1000")
            {
                testSelButton0.Visible = true;
                testSelButton1.Visible = true;
                testSelButton2.Visible = true;
                testSelButton3.Visible = true;
                testSelButton8.Visible = true;
                testSelButton4.Visible = true;
                testSelButton5.Visible = true;
                testSelButton6.Visible = true;
                testSelButton7.Visible = true;
                AUXCableOption.Visible = true;
                if (AUXCableOption.Checked == false)
                {
                    testSelButton9.Visible = true;
                }

                testSelButton0.Text = "MP3 Input Sensitivity / Max. SPL -------------------------------------------------------------------------------------------";
                testSelButton1.Text = "Volume Control Knob Function -----------------------------------------------------------------------------------------------";
                testSelButton2.Text = "Microphone Input Sensitivity -----------------------------------------------------------------------------------------------";
                testSelButton3.Text = "Audio Effect /w Sound=Wide, VB=ON -------------------------------------------------------------------------------------------";
                testSelButton4.Text = "Audio Effect /w Sound=Narrow, VB=OFF ----------------------------------------------------------------------------------------";
                testSelButton5.Text = "Audio Effect /w Sound=Narrow, VB=ON -----------------------------------------------------------------------------------------";
                testSelButton6.Text = "System Background Noise Output ----------------------------------------------------------------------------------------------";
                testSelButton7.Text = "Limited Key Switch Function -------------------------------------------------------------------------------------------------";
                testSelButton8.Text = "Self Test -------------------------------------------------------------------------------------------------------------------";
                //      testSelButton9.Text = "AUX Cable -------------------------------------------------------------------------------------------------------------------";
                if (AUXCableOption.Checked == false)
                {
                    testSelButton9.Text = "AUX Cable -------------------------------------------------------------------------------------------------------------------";
                }

                testResult0.ForeColor = Color.LightGray;
                testResult0.ForeColor = Color.LightGray;
                testResult0.Text = "NT";
                testResult1.ForeColor = Color.LightGray;
                testResult1.ForeColor = Color.LightGray;
                testResult1.Text = "NT";
                testResult2.ForeColor = Color.LightGray;
                testResult2.ForeColor = Color.LightGray;
                testResult2.Text = "NT";
                testResult3.ForeColor = Color.LightGray;
                testResult3.ForeColor = Color.LightGray;
                testResult3.Text = "NT";
                testResult4.ForeColor = Color.LightGray;
                testResult4.ForeColor = Color.LightGray;
                testResult4.Text = "NT";
                testResult5.ForeColor = Color.LightGray;
                testResult5.ForeColor = Color.LightGray;
                testResult5.Text = "NT";
                testResult6.ForeColor = Color.LightGray;
                testResult6.ForeColor = Color.LightGray;
                testResult6.Text = "NT";
                testResult7.ForeColor = Color.LightGray;
                testResult7.ForeColor = Color.LightGray;
                testResult7.Text = "NT";
                testResult8.ForeColor = Color.LightGray;
                testResult8.ForeColor = Color.LightGray;
                testResult8.Text = "NT";
                //  testResult9.ForeColor = Color.LightGray; testResult9.ForeColor = Color.LightGray; testResult9.Text = "NT";
                if (AUXCableOption.Checked == false)
                {
                    testResult9.ForeColor = Color.LightGray;
                    testResult9.ForeColor = Color.LightGray;
                    testResult9.Text = "NT";
                }
            }
            else if (DUT.productModel == "SS400")
            {
                testSelButton0.Visible = true;
                testSelButton1.Visible = true;
                testSelButton0.Text = "System Sensitivity / Max. SPL ---------------------------------------------------------------------------------";
                testSelButton1.Text = "System Functionality Check /w Accelerometer -------------------------------------------------------------------";

                testResult0.ForeColor = Color.LightGray;
                testResult0.ForeColor = Color.LightGray;
                testResult0.Text = "NT";
                testResult1.ForeColor = Color.LightGray;
                testResult1.ForeColor = Color.LightGray;
                testResult1.Text = "NT";
            }
            else if (productButton18.Checked)
            {
                testSelButton0.Visible = true;
                testSelButton1.Visible = true;
                testSelButton6.Visible = true;
                testSelButton0.Text = "Frequency Sweep -------------------------------------------------------------------------------------------";
                testSelButton1.Text = "Maximum SPL -----------------------------------------------------------------------------------------------";
                testSelButton6.Text = "Maximum Output Noise Level --------------------------------------------------------------------------------";

                testResult0.ForeColor = Color.LightGray;
                testResult0.ForeColor = Color.LightGray;
                testResult0.Text = "NT";
                testResult1.ForeColor = Color.LightGray;
                testResult1.ForeColor = Color.LightGray;
                testResult1.Text = "NT";
                testResult6.ForeColor = Color.LightGray;
                testResult6.ForeColor = Color.LightGray;
                testResult6.Text = "NT";

                if (DUT.productModel != "950NXT")
                {
                    testSelButton9.Visible = true;
                    testSelButton9.Text = "Cooling Fans Function ----------------------------------------------------------------------------------";
                    testResult9.ForeColor = Color.LightGray;
                    testResult9.ForeColor = Color.LightGray;
                    testResult9.Text = "NT";
                }
            }
            else
            {
                testSelButton0.Visible = true;
                testSelButton0.Text = "System Sensitivity / Max. SPL -------------------------------------------------------------------";
                testResult0.ForeColor = Color.LightGray;
                testResult0.ForeColor = Color.LightGray;
                testResult0.Text = "NT";
            }

        } //*** end of setupTestSelection()

        void displaySetupImages()
        {
            string setupImageName = null, prodModel = null;

            pictureBox1.Image = null;
            pictureBox2.Image = null;
            pictureBox3.Image = null;
            pictureBox4.Image = null;
            pictureBox5.Image = null;
            pictureBox6.Image = null;
            pictureBox7.Image = null;
            pictureBox8.Image = null;
            pictureBox9.Image = null;
            pictureBox10.Image = null;
            pictureBox11.Image = null;
            pictureBox12.Image = null;

            textBoxForSetupPic1.Text = textBoxForSetupPic2.Text = textBoxForSetupPic3.Text = null;
            textBoxForSetupPic4.Text = textBoxForSetupPic5.Text = textBoxForSetupPic6.Text = null;
            textBoxForSetupPic7.Text = textBoxForSetupPic8.Text = textBoxForSetupPic9.Text = null;
            textBoxForSetupPic10.Text = textBoxForSetupPic11.Text = textBoxForSetupPic12.Text = null;

            if ((productButton12.Checked) || (productButton13.Checked)) prodModel = "60-DEG";
            else if (productButton10.Checked) prodModel = "360X";
            else if (productButton14.Checked) prodModel = "60-DEG wo Amp";
            else if (productButton22.Checked) prodModel = "360XL-MID";
            else prodModel = DUT.productModel;

            for (int cnt = 1; cnt < 13; cnt++)
            {
                setupImageName = prodModel + "_" + cnt.ToString() + ".jpg";
                if (File.Exists(Folder.setupPicsFolder + setupImageName))
                {
                    switch (cnt)
                    {
                        case 1:
                            pictureBox1.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic1.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption1") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic1.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption1_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        case 2:
                            pictureBox2.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic2.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption2") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic2.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption2_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        case 3:
                            pictureBox3.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic3.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption3") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic3.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption3_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        case 4:
                            pictureBox4.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic4.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption4") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic4.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption4_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        case 5:
                            pictureBox5.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic5.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption5") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic5.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption5_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        case 6:
                            pictureBox6.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic6.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption6") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic6.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption6_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        case 7:
                            pictureBox7.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic7.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption7") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic7.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption7_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        case 8:
                            pictureBox8.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic8.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption8") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic8.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption8_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        case 9:
                            pictureBox9.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic9.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption9") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic9.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption9_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        case 10:
                            pictureBox10.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic10.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption10") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic10.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption10_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        case 11:
                            pictureBox11.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic11.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption11") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic11.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption11_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        case 12:
                            pictureBox12.Image = Image.FromFile(Folder.setupPicsFolder + setupImageName);
                            textBoxForSetupPic12.Text = ini.IniReadValue(DUT.productModel, "SetupPictureCaption12") + "\r\n";
                            for (int x = 1; x < 4; x++)
                            {
                                textBoxForSetupPic12.Text += ini.IniReadValue(DUT.productModel, "SetupPictureCaption12_" + Convert.ToString(x)) + "\r\n";
                            }

                            break;
                        default:
                            break;
                    }
                }
            }

        } //*** end of displaySetupImages()

        void setupSerialNumberPerProduct()
        {
            SN1_Label.Visible = true;
            SN2_Label.Visible = true;
            SN3_Label.Visible = true;
            SN4_Label.Visible = true;
            SN5_Label.Visible = false;
            SN6_Label.Visible = false;
            SN7_Label.Visible = false;
            SN8_Label.Visible = false;
            SN9_Label.Visible = false;
            SN10_Label.Visible = false;

            SN1_Label.Text = "System";
            SN2_Label.Text = "Control Unit";
            SN3_Label.Text = "Electronics";
            SN4_Label.Text = "Microphone";

            SN_Box1.Visible = true;
            SN_Box2.Visible = true;
            SN_Box3.Visible = true;
            SN_Box4.Visible = true;
            SN_Box10.Visible = false;
            SN_Box5.Visible = false;
            SN_Box6.Visible = false;
            SN_Box7.Visible = false;
            SN_Box8.Visible = false;
            SN_Box9.Visible = false;
            SN_Box10.Visible = false;
            SN_Box2.Enabled = SN_Box3.Enabled = SN_Box4.Enabled = SN_Box5.Enabled = SN_Box6.Enabled = SN_Box7.Enabled = SN_Box8.Enabled = SN_Box9.Enabled = SN_Box10.Enabled = true;

            SN_Box1.Text = null;
            SN_Box2.Text = null;
            SN_Box3.Text = null;
            SN_Box4.Text = null;
            SN_Box5.Text = null;
            SN_Box6.Text = SN_Box7.Text = SN_Box8.Text = SN_Box9.Text = SN_Box10.Text = null;

            //100Xs, 500Xs, 1000X, 1000XVB
            if ((productButton0.Checked) || (DUT.productModel == "500X") || (productButton5.Checked) || (productButton8.Checked))
            {
                SN2_Label.Text = "MP3 Player";

                if ((productButton0.Checked))
                {
                    SN_Box5.Visible = true;
                    SN5_Label.Visible = true;
                    SN5_Label.Text = "Battery";
                    if (DUT.productModel == "100X-NAVY")
                    {
                        SN2_Label.Visible = false;
                        SN_Box2.Visible = false;
                    }
                }
                else if ((productButton5.Checked) || (productButton8.Checked))
                {
                    SN3_Label.Text = "Amp. Pack";
                }

            }
            else if (DUT.productModel == "1000X2U")
            {
                SN2_Label.Visible = false;
                SN_Box2.Visible = false;
                SN4_Label.Visible = false;
                SN_Box4.Visible = false;
                SN3_Label.Text = "2U Chassis";
            }
            else if (DUT.productModel == "1000Xi")
            {
                SN5_Label.Visible = true;
                SN_Box5.Visible = true;
                SN5_Label.Text = "CarbonFiber Hd";
            }
            else if (DUT.productModel == "1000")
            {
                SN5_Label.Visible = true;
                SN6_Label.Visible = true;
                SN7_Label.Visible = true;
                SN_Box5.Visible = true;
                SN_Box6.Visible = true;
                SN_Box7.Visible = true;
                SN1_Label.Text = "1000-AHD-G";
                SN2_Label.Text = "X-MP3-AL";
                SN3_Label.Text = "1000-AMP";
                SN4_Label.Text = "X-MIC-AL";
                SN5_Label.Text = "1000-G-SYS";
                SN6_Label.Text = "1000-AC-PWR";
                SN7_Label.Text = "PHRASLTR-2";
            }
            //360X Manifolds, DS60 horns only & SSX products
            else if ((productButton10.Checked) || (productButton14.Checked) || (productButton17.Checked))
            {
                SN2_Label.Visible = false;
                SN3_Label.Visible = false;
                SN4_Label.Visible = false;
                SN_Box2.Visible = false;
                SN_Box3.Visible = false;
                SN_Box4.Visible = false;
            }
            else if ((productButton12.Checked) || (productButton13.Checked)) //60DS w Tranfs
            {
                SN3_Label.Visible = false;
                SN4_Label.Visible = false;
                SN_Box3.Visible = false;
                SN_Box4.Visible = false;
                if (productButton15.Checked) SN2_Label.Text = "Driver";
                else SN2_Label.Text = "Transf. Box";
            }
            else if (productButton15.Checked) //SS
            {
                if (DUT.productModel == "SS400")
                {
                    SN3_Label.Visible = false;
                    SN4_Label.Visible = false;
                    SN_Box3.Visible = false;
                    SN_Box4.Visible = false;
                }
                else //SS100/SS300
                {
                    SN4_Label.Visible = false;
                    SN_Box4.Visible = false;
                    SN3_Label.Text = "Transformer";
                }

                SN2_Label.Text = "Driver Panel";

            }
            else if (productButton16.Checked) //SSX /w Trans
            {
                SN3_Label.Visible = SN4_Label.Visible = false;
                SN_Box3.Visible = SN_Box4.Visible = false;

                SN2_Label.Text = "Transformer";
            }
            else if (productButton18.Checked) //RXs
            {
                SN5_Label.Visible = true;
                SN6_Label.Visible = true;
                SN7_Label.Visible = true;
                SN8_Label.Visible = true;
                SN9_Label.Visible = true;
                SN10_Label.Visible = true;
                SN_Box5.Visible = true;
                SN_Box6.Visible = true;
                SN_Box7.Visible = true;
                SN_Box8.Visible = true;
                SN_Box9.Visible = true;
                SN_Box9.Visible = true;
                SN_Box10.Visible = true;
                SN2_Label.Text = "Firmware Ver";
                SN3_Label.Text = "Electronics";
                SN4_Label.Text = "Head No 1";
                SN5_Label.Text = "Head No 2";
                SN6_Label.Text = "AMC No 1";
                SN7_Label.Text = "AMC No 2";
                SN8_Label.Text = "Light";
                SN9_Label.Text = "Camera";
                SN10_Label.Text = "48VDC Ext PS";

                if (DUT.productModel == "1000RX")
                {
                    SN_Box5.Text = "N/A";
                    SN_Box5.Enabled = false;
                }
                else
                {
                    SN_Box5.Text = "";
                    SN_Box5.Enabled = true;
                    if (DUT.productModel == "950NXT")
                    {
                        SN6_Label.Text = "K300 PC";
                        SN7_Label.Text = "Main PWR PCA";
                    }
                }
            }
            else if ((productButton20.Checked) || (productButton21.Checked)) //360Xm /w & /wo Trans
            {
                SN_Box2.Visible = SN_Box3.Visible = SN_Box4.Visible = false;
                SN2_Label.Visible = SN3_Label.Visible = SN4_Label.Visible = false;
            }
            else if (productButton22.Checked) //360XL-MID
            {
                SN_Box2.Visible = SN_Box3.Visible = SN_Box4.Visible = false;
                SN2_Label.Visible = SN3_Label.Visible = SN4_Label.Visible = false;
            }
        }

        void setupDriverSerialNumberPerProduct()
        {
            driverSN_Box1.Visible = false;
            driverSN_Box2.Visible = false;
            driverSN_Box3.Visible = false;
            driverSN_Box4.Visible = false;
            driverSN_Box5.Visible = false;
            driverSN_Box6.Visible = false;
            driverSN_Box7.Visible = false;
            driverSN_Box8.Visible = false;

            if ((productButton0.Checked) || (productButton11.Checked) || (productButton12.Checked) || (productButton13.Checked) || (productButton14.Checked))
            {
                driverSN_Box1.Visible = true;
            }
            else if ((productButton1.Checked) || (productButton2.Checked))
            {
                driverSN_Box1.Visible = true;
                driverSN_Box2.Visible = true;
            }
            else if ((productButton3.Checked) || (productButton4.Checked) || (productButton10.Checked) || (DUT.productModel == "500RX") || (DUT.productModel == "950RXL") || (DUT.productModel == "950NXT"))
            {
                driverSN_Box1.Visible = true;
                driverSN_Box2.Visible = true;
                driverSN_Box3.Visible = true;
                driverSN_Box4.Visible = true;
            }
            else if ((productButton5.Checked) || (productButton6.Checked) || (productButton7.Checked) || (productButton8.Checked) || (productButton9.Checked) || (DUT.productModel == "1000RX"))
            {
                driverSN_Box1.Visible = true;
                driverSN_Box2.Visible = true;
                driverSN_Box3.Visible = true;
                driverSN_Box4.Visible = true;
                driverSN_Box5.Visible = true;
                driverSN_Box6.Visible = true;
                driverSN_Box7.Visible = true;
            }
            else if (productButton19.Checked)
            {
                driverSN_Box1.Visible = true;
                driverSN_Box2.Visible = true;
                driverSN_Box3.Visible = true;
                driverSN_Box4.Visible = true;
                driverSN_Box5.Visible = true;
                driverSN_Box6.Visible = true;
                driverSN_Box7.Visible = true;
                driverSN_Box8.Visible = true;
            }
            else if (productButton20.Checked)
            {
                if ((modelButton0.Checked) || (modelButton1.Checked))
                {
                    driverSN_Box1.Visible = true;
                }
                else if ((modelButton2.Checked) || (modelButton3.Checked) || (modelButton4.Checked))
                {
                    driverSN_Box1.Visible = driverSN_Box2.Visible = true;
                }
                else
                {
                    driverSN_Box1.Visible = driverSN_Box2.Visible = driverSN_Box3.Visible = driverSN_Box4.Visible = true;
                }
            }
            else if (productButton21.Checked)
            {
                if (modelButton0.Checked)
                {
                    driverSN_Box1.Visible = true;
                }
                else
                {
                    driverSN_Box1.Visible = driverSN_Box2.Visible = true;
                }
            }
            else if (productButton22.Checked)
            {
                driverSN_Box1.Visible = true;
                driverSN_Box2.Visible = true;

                if (DUT.productModel == "360XL-MID 2-ST")
                {
                    driverSN_Box2.Text = "";
                }
                else
                {
                    driverSN_Box2.Text = "NA";
                }
            }
            else { }

        }

        private void startButton_Click(object sender, EventArgs e)
        {
            if (DUT.Debug == 0)
            {
                if (MissingSN()) return;
            }

            shutdownTimer.Stop();
            if (DUT.autoShutdown == 1)
            {
                shutdownTimer.Start();
            }

            date_Stamp.Text = DateTime.Now.ToString("MM/dd/yyyy");
            time_Stamp.Text = DateTime.Now.ToString("hh:mm tt");

            if (troubleShootingCheckBox.Checked)
            {
                recordSampleInfo(true);
            }
            else
            {
                recordSampleInfo(false);
            }

            //*** check for retested S/Ns ***// 
            if (DUT.Debug == 0)
            {
                if (!(troubleShootingCheckBox.Checked) && !(DUT.productModel == "450XL Extended Test") && !(preTest.Checked) && !(repaired.Checked) && (DuplicatedSNWasFound())) return;
            }

            String TestTimeStamp = DateTime.Now.ToString("HH") + DateTime.Now.ToString("mm") + DateTime.Now.ToString("ss");

            De_activateUserInputs(true);
            string FileNameForFR = assignFileNameForFR(); //step2
            string FileNameForMaxOutput = assignFileNameForMaxOutput();

            DUT.hasResults = true;

            SetupSwitchBox();

            if (testSelButton0.Checked) //test0
            {
                GetTestParameters("stimulusFR", "stimulusSPL", "sweepLowerLimit", "sweepUpperLimit", "WindowBegin", "WindowEnd", "FreqSPL");

                if (productButton18.Checked) //RXs
                {
                    if (SetupSequenceSelect("MP3_Ch1", "FreqSweep", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                    {
                        De_activateUserInputs(false);
                        return;
                    };
                    runSteppedFreqSweep(FileNameForFR);
                    DisplayTestResults("FR");
                    De_activateUserInputs(false);
                    return;
                }

                /******************************************************************************************************************************************************************/
                if (SetupSequenceSelect("MP3_Ch1", "FR", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                {
                    De_activateUserInputs(false);
                    return;
                };

                runFRTest(FileNameForFR, null);
                DisplayTestResults("FR");

                if (SetupSequenceSelect("MP3_Ch1", "SPL", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                {
                    De_activateUserInputs(false);
                    return;
                };
                runMaxSPLTest(0, false, FileNameForMaxOutput);

                DisplayTestResults("SPL");

            }
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            else if (testSelButton1.Checked) //test1
            {
                //   testResult1.Text = "NT         "; testResult1.Update();               
                //    testResult1.ForeColor = Color.LightGray; testResult1.Text = "NT"; testResult1.Update();

                GetTestParameters("stimulusFR1", "stimulusSPL1", "sweepLowerLimit1", "sweepUpperLimit1", "WindowBegin", "WindowEnd", "FreqSPL1");

                if (productButton18.Checked) //RXs
                {
                    if (SetupSequenceSelect("MP3_Ch1", "SPL", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                    {
                        De_activateUserInputs(false);
                        return;
                    };
                    runMaxSPLTest(0, false, FileNameForMaxOutput);
                    DisplayTestResults("SPL");

                    De_activateUserInputs(false);
                    return;
                }
                else if ((productButton12.Checked) || (productButton13.Checked) || (productButton15.Checked) || (productButton16.Checked) || (productButton20.Checked))
                {
                    if (DUT.productModel == "SS400")
                    {
                        if (SetupSequenceSelect("MP3_Ch1", "SPL", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                        {
                            De_activateUserInputs(false);
                            return;
                        };
                        runMaxSPLTest(1, true, FileNameForMaxOutput); //recording Mic level
                        DisplayTestResults("SPL");

                        De_activateUserInputs(false);
                        return;
                    }

                    if ((testResult0.Text == "FAILED") || (testResult0.Text == "NT"))
                    {
                        DialogResult dialogResult = MessageBox.Show("HIGH Power test needed to be performed with good result first.\nDo you want to continue with MID Power test?\n\n",
                          "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (dialogResult == DialogResult.No) //
                        {
                            channelSetup(0x00, 0x00, DUT.portCState);
                            reportPassFail1.Text = "  N/T  ";
                            //    testResult1.Text = "NT      ";
                            De_activateUserInputs(false);
                            return;
                        }
                    }
                }

                if (SetupSequenceSelect("MP3_Ch1", "FR", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                {
                    De_activateUserInputs(false);
                    return;
                };

                runFRTest(FileNameForFR, null);
                DisplayTestResults("FR");

                if (DUT.Debug == 1)
                {
                    debugLabel6.Text = "Test 0 sensitivity=" + DUT.PFForTest0Sel[0] + ", SPL=" + DUT.PFForTest0Sel[1];
                }

            }
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            else if (testSelButton2.Checked) //test2
            {
                //     testResult2.ForeColor = Color.LightGray; testResult2.Text = "NT     "; testResult2.Update();            
                if (productButton18.Checked) //Alert Tone SPL in RXs
                {
                    if (SetupSequenceSelect("MP3_Ch1", "RX_SPL", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                    {
                        De_activateUserInputs(false);
                        return;
                    };
                    runMaxSPLTest(2, false, FileNameForMaxOutput);
                    DisplayTestResults("SPL");

                    De_activateUserInputs(false);
                    return;
                }

                GetTestParameters("stimulusFR2", "stimulusSPL", "sweepLowerLimit2", "sweepUpperLimit2", "WindowBegin", "WindowEnd", "FreqSPL");

                if ((DUT.noBlackBoxSetup == 1) && !((productButton6.Checked) || (productButton10.Checked) || (productButton12.Checked) || (productButton13.Checked) || (productButton14.Checked) ||
                    (productButton15.Checked) || (productButton16.Checked) || (productButton17.Checked) || (productButton18.Checked) || (productButton19.Checked) ||
                    (productButton20.Checked) || (productButton21.Checked) || (productButton22.Checked)))
                {
                    if (SetupSequenceSelect("MIC_Ch2", "FR", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                    {
                        De_activateUserInputs(false);
                        return;
                    }; //AP output CH2 --> MIC
                }
                else
                {
                    if (SetupSequenceSelect("MP3_Ch1", "FR", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                    {
                        De_activateUserInputs(false);
                        return;
                    }; //AP output CH1 --> MP3
                }

                if ((productButton12.Checked) || (productButton13.Checked) || (productButton15.Checked) || (productButton16.Checked) || (productButton20.Checked))
                {
                    if ((testResult0.Text == "FAILED") || (testResult0.Text == ""))
                    {
                        DialogResult dialogResult = MessageBox.Show("HIGH Power test needed to be performed with good result first.\nDo you want to continue with LOW POWER test?\n\n",
                          "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (dialogResult == DialogResult.No) //
                        {
                            channelSetup(0x00, 0x00, DUT.portCState);
                            reportPassFail2.Text = "  N/T  ";
                            //      testResult2.Text = "NT      ";
                            De_activateUserInputs(false);
                            return;
                        }
                    }
                }

                runFRTest(FileNameForFR, null);
                DisplayTestResults("FR");

            }
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            else if (testSelButton3.Checked) //test3
            {
                //    testResult3.ForeColor = Color.LightGray; testResult3.Text = "NT     "; testResult3.Update();
                GetTestParameters("stimulusFR", "stimulusSPL", "sweepLowerLimit3", "sweepUpperLimit3", "WindowBegin3", "WindowEnd3", "FreqSPL");
                if (SetupSequenceSelect("MP3_Ch1", "FR", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                {
                    De_activateUserInputs(false);
                    return;
                };
                runFRTest(FileNameForFR, null);
                DisplayTestResults("FR");

                if (productButton20.Checked)
                {
                    if (SetupSequenceSelect("MP3_Ch1", "SPL", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                    {
                        De_activateUserInputs(false);
                        return;
                    };
                    runMaxSPLTest(3, false, FileNameForMaxOutput);

                    DisplayTestResults("SPL");

                }
            }
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            else if (testSelButton4.Checked) //test4
            {
                //    testResult4.ForeColor = Color.LightGray; testResult4.Text = "NT     "; testResult4.Update();
                GetTestParameters("stimulusFR", "stimulusSPL", "sweepLowerLimit4", "sweepUpperLimit4", "WindowBegin4", "WindowEnd4", "FreqSPL");
                if (SetupSequenceSelect("MP3_Ch1", "FR", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                {
                    De_activateUserInputs(false);
                    return;
                };

                if (productButton20.Checked)
                {
                    if ((testResult3.Text == "FAILED") || (testResult3.Text == "NT"))
                    {
                        DialogResult dialogResult = MessageBox.Show("HIGH Power test needed to be performed with good result first.\nDo you want to continue with MID Power test?\n\n",
                          "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (dialogResult == DialogResult.No) //
                        {
                            channelSetup(0x00, 0x00, DUT.portCState);
                            reportPassFail1.Text = "  N/T  ";
                            //    testResult1.Text = "NT      ";
                            De_activateUserInputs(false);
                            return;
                        }
                    }
                }

                runFRTest(FileNameForFR, null);
                DisplayTestResults("FR");

            }
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            else if (testSelButton5.Checked) //test5
            {
                //    testResult5.ForeColor = Color.LightGray; testResult5.Text = "NT     "; testResult5.Update();
                GetTestParameters("stimulusFR", "stimulusSPL", "sweepLowerLimit5", "sweepUpperLimit5", "WindowBegin5", "WindowEnd5", "FreqSPL");
                if (SetupSequenceSelect("MP3_Ch1", "FR", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                {
                    De_activateUserInputs(false);
                    return;
                };

                runFRTest(FileNameForFR, null);
                DisplayTestResults("FR");

            }
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            else if (testSelButton6.Checked) //test6
            {
                //    testResult6.ForeColor = Color.LightGray; testResult6.Text = ""; testResult6.Update();
                GetTestParameters("stimulusFR6", "stimulusSPL6", "sweepLowerLimit6", "sweepUpperLimit6", "WindowBegin", "WindowEnd", "FreqSPL");
                if (SetupSequenceSelect("MP3_Ch1", "Noise", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                {
                    De_activateUserInputs(false);
                    return;
                };

                runNoiseTest(1, FileNameForFR);
                DisplayTestResults("FR");

            }
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            else if (testSelButton7.Checked) //test7
            {
                //    testResult7.ForeColor = Color.LightGray; testResult7.Text = ""; testResult7.Update();
                GetTestParameters("stimulusFR", "stimulusSPL", "sweepLowerLimit7", "sweepUpperLimit7", "WindowBegin", "WindowEnd", "FreqSPL");
                if (SetupSequenceSelect("MP3_Ch1", "FR", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                {
                    De_activateUserInputs(false);
                    return;
                };

                runFRTest(FileNameForFR, null);
                DisplayTestResults("FR");

                if (DUT.productModel == "100X-NAVY")
                {
                    if (SetupSequenceSelect("MP3_Ch1", "SPL", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                    {
                        De_activateUserInputs(false);
                        return;
                    };
                    runMaxSPLTest(0, false, FileNameForMaxOutput);

                    DisplayTestResults("SPL");
                }

            }
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            else if (testSelButton8.Checked) //test8
            {
                //     testResult8.ForeColor = Color.LightGray; testResult8.Text = ""; testResult8.Update();

                DialogResult dialogResult = MessageBox.Show("(1) Remove all test cable from MIC input\n" +
                  "(2) Connect MP3 Player to test unit (MP3 INPUT)\n" +
                  "(3) Turn Volume Knob to Maximum position, Sound=Wide, VB=Off, Limited Key SW=Max.\n" +
                  "(4) Press the TONE button ONCE, then press the TEST button (Within 5 seconds)\n" +
                  "(5) Wait for the Self Test to stop by itself\n" +
                  "(6) Press the TONE button again to complete the Self Test\n\n" +
                  "Does the LED indicator show GREEN?  \n", "Confirm", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (dialogResult == DialogResult.Yes)
                {
                    testResult8.ForeColor = Color.ForestGreen;
                    testResult8.Text = "PASSED";
                }
                else if (dialogResult == DialogResult.No)
                {
                    testResult8.ForeColor = Color.Red;
                    testResult8.Text = "FAILED";
                }
                /*      else
                      {
                          testResult8.ForeColor = Color.LightGray; testResult8.Text = ""; testResult8.Update();                   
                      } */

                DisplayTestResults("FR");
            }
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            else if (testSelButton9.Checked) //test9
            {
                //     testResult9.ForeColor = Color.LightGray; testResult9.Text = ""; testResult9.Update();
                if (productButton18.Checked) //If RX's
                {
                    DialogResult dialogResult = MessageBox.Show("From LRAD Controller GUI, manually toggle cooling fans ON & OFF as shown below\n" +
                      "Select button Tools --> Debug Tab --> Enter Command 020102010100 in ''Pan-Tilt'' section\n" +
                      "and press ''Send'' to turn ON cooling Fans\n" +
                      "Observe the colling fans activity\n\n" +
                      "Are the cooling fans RUNNING?", "Confirm", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (dialogResult == DialogResult.Yes)
                    {
                        DialogResult dialogResult1 = MessageBox.Show("Now turn OFF the cooling fans as shown below\n" +
                          "Tools --> Debug --> Enter Command 020102010000\n" +
                          "and press ''Send'' to turn OFF cooling Fans\n" +
                          "Observe the colling fans activity\n\n" +
                          "Are the cooling fans STOPPED?", "Confirm", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                        if (dialogResult1 == DialogResult.Yes)
                        {
                            testResult9.ForeColor = Color.ForestGreen;
                            testResult9.Text = "PASSED";
                        }
                        else if (dialogResult1 == DialogResult.No)
                        {
                            testResult9.ForeColor = Color.Red;
                            testResult9.Text = "FAILED";
                        }
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        testResult9.ForeColor = Color.Red;
                        testResult9.Text = "FAILED";
                    }

                    DisplayTestResults("FR");
                    De_activateUserInputs(false);

                    return;
                }

                GetTestParameters("stimulusFR", "stimulusSPL", "sweepLowerLimit9", "sweepUpperLimit9", "WindowBegin", "WindowEnd", "FreqSPL");
                if (SetupSequenceSelect("MP3_Ch1", "FR", GetRefFileName(), GetRefFileNamePlus(), GetRefFileNameMinus()))
                {
                    De_activateUserInputs(false);
                    return;
                };

                DUT.FRTestStatus2 = null; //Result for Ring signal 

                runFRTest(FileNameForFR, "Tip"); //AUX cable Tip sig

                if (troubleShootingCheckBox.Checked)
                {
                    FileNameForFR = "\\Screening Test 9 (Ring) - ";
                }
                else
                {
                    FileNameForFR = "\\[" + DUT.SN[1] + "] AUX Cable (Ring) - ";
                }

                Thread.Sleep(500);
                Set_LRADX_MIC_LowerMic();
                Thread.Sleep(500);

                runFRTest(FileNameForFR, "Ring"); //AUX cable Ring sig

                DisplayTestResults("FR");
            }

            channelSetup(0x00, 0x00, DUT.portCState);
            De_activateUserInputs(false);
        }

        /****************************************************************************************************************/
        /* Display the test results & Final report                                                                      */
        /****************************************************************************************************************/
        private void DisplayTestResults(string ResultType)
        {
            if (testSelButton0.Checked)
            {
                if (productButton18.Checked) //RXs
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        reportPassFail0.Text = DUT.FRTestStatus;
                    }

                    // reportPassFail0.Text = "PASSED"; //SPDebug

                    testResult0.Text = DUT.FRTestStatus;

                    if (DUT.FRTestStatus == "FAILED")
                    {
                        testResult0.ForeColor = Color.Red;
                    }
                    else
                    {
                        testResult0.ForeColor = Color.LightGreen;
                    }
                }
                else
                {
                    if (ResultType == "FR")
                    {
                        if (!troubleShootingCheckBox.Checked)
                        {
                            reportPassFail0.Text = DUT.FRTestStatus;
                        }
                    }
                    else if (ResultType == "SPL")
                    {
                        if (((productButton12.Checked) || (productButton13.Checked) || (productButton15.Checked) || (productButton16.Checked)) && (!troubleShootingCheckBox.Checked)) //models w Trans
                        {
                            if (DUT.productModel == "360Xm 4-ST 100V-240W")
                            {
                                reportPassFail1.Text = DUT.SPLTestStatus;
                            }
                            else if (DUT.productModel != "SS400")
                            {
                                reportPassFail3.Text = DUT.SPLTestStatus;
                            }
                        }
                        else if ((DUT.productModel == "360Xm 1-ST 25V-60W") || (DUT.productModel == "360Xm 1-ST 100V-60W") || (DUT.productModel == "360Xm 2-ST 100V-120W"))
                        {
                            reportPassFail3.Text = DUT.SPLTestStatus;
                        }
                        else if ((DUT.productModel == "360Xm 4-ST 100V-240W"))
                        {
                            reportPassFail1.Text = DUT.SPLTestStatus;
                        }
                        else if ((productButton10.Checked) || (productButton14.Checked) || (productButton17.Checked) || (productButton22.Checked)) //360X manifolds, SSX wo Trans & 360X-MIDs
                        {
                            if (!troubleShootingCheckBox.Checked)
                            {
                                reportPassFail1.Text = DUT.SPLTestStatus;
                            }
                        }

                        if ((DUT.FRTestStatus == "FAILED") || (DUT.SPLTestStatus == "FAILED"))
                        {
                            testResult0.ForeColor = Color.Red;
                            testResult0.Text = "FAILED"; //to define FR+SPL pass/fail 

                        }
                        else
                        {
                            testResult0.ForeColor = Color.LightGreen;
                            testResult0.Text = "PASSED";
                        }
                    }
                }
            }
            else if (testSelButton1.Checked)
            {
                if ((productButton18.Checked) || (DUT.productModel == "SS400")) //RXs & SS400
                {
                    testResult1.Text = DUT.PFPerTestItem[1] = DUT.SPLTestStatus;

                    if (!troubleShootingCheckBox.Checked)
                    {
                        reportPassFail1.Text = DUT.SPLTestStatus;
                    }

                    //   reportPassFail1.Text = "PASSED"; //SPDebug

                    if (DUT.SPLTestStatus == "FAILED")
                    {
                        testResult1.ForeColor = Color.Red;
                    }
                    else
                    {
                        testResult1.ForeColor = Color.LightGreen;
                    }
                }
                else
                {
                    testResult1.Text = DUT.PFPerTestItem[1] = DUT.FRTestStatus;

                    if (!troubleShootingCheckBox.Checked)
                    {
                        reportPassFail1.Text = DUT.FRTestStatus;
                    }

                    if (DUT.FRTestStatus == "FAILED")
                    {
                        testResult1.ForeColor = Color.Red;
                    }
                    else
                    {
                        testResult1.ForeColor = Color.LightGreen;
                    }
                }

                /**************************************************************************************
                        reportPassFail0.Text = "PASSED"; //Debug
                        reportPassFail1.Text = "PASSED"; //Debug
                     //   reportPassFail2.Text = "PASSED"; //Debug
                        reportPassFail3.Text = "PASSED"; //Debug
                     //   reportPassFail6.Text = "PASSED"; //Debug

                        /**************************************************************************************/
            }
            else if (testSelButton2.Checked)
            {
                testResult2.Text = DUT.PFPerTestItem[2] = DUT.FRTestStatus;

                if (!troubleShootingCheckBox.Checked)
                {
                    reportPassFail2.Text = DUT.FRTestStatus;
                }

                if (DUT.FRTestStatus == "FAILED")
                {
                    testResult2.ForeColor = Color.Red;
                }
                else
                {
                    testResult2.ForeColor = Color.LightGreen;
                }
            }
            else if (testSelButton3.Checked)
            {
                testResult3.Text = DUT.PFPerTestItem[3] = DUT.FRTestStatus;

                if (DUT.FRTestStatus == "FAILED")
                {
                    testResult3.ForeColor = Color.Red;
                }
                else
                {
                    testResult3.ForeColor = Color.LightGreen;
                }

                if ((DUT.productModel == "100X-NAVY") || (productButton8.Checked)) //1000VB's
                                                                                   //        if ((DUT.productModel == "100X-NAVY") || (DUT.productModel == "1000XVB"))
                {
                    if ((testResult3.Text == "FAILED") || (testResult4.Text == "FAILED"))
                    {
                        if (!troubleShootingCheckBox.Checked)
                        {
                            reportPassFail3.Text = "FAILED";
                        }
                    }
                    else if ((testResult3.Text == "PASSED") && (testResult4.Text == "PASSED"))
                    {
                        if (!troubleShootingCheckBox.Checked)
                        {
                            reportPassFail3.Text = "PASSED";
                        }
                    }
                    else
                    {
                        reportPassFail3.Text = "NC";
                    }
                }
                else if (DUT.productModel == "1000X")
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        reportPassFail3.Text = DUT.FRTestStatus;
                    }
                }
                else if ((productButton20.Checked) && (ResultType == "SPL"))
                {
                    if ((DUT.FRTestStatus == "FAILED") || (DUT.SPLTestStatus == "FAILED"))
                    {
                        testResult3.ForeColor = Color.Red;
                        reportPassFail3.Text = testResult3.Text = "FAILED"; //to define FR+SPL pass/fail                      
                    }
                    else
                    {
                        testResult3.ForeColor = Color.LightGreen;
                        reportPassFail3.Text = "PASSED";
                    }
                }
                else
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        if ((testResult3.Text == "FAILED") || (testResult4.Text == "FAILED") || (testResult5.Text == "FAILED")) reportPassFail3.Text = "FAILED";
                        else if ((testResult3.Text == "PASSED") && (testResult4.Text == "PASSED") && (testResult5.Text == "PASSED")) reportPassFail3.Text = "PASSED";
                    }
                    else reportPassFail3.Text = "NC";
                }

            }
            else if (testSelButton4.Checked)
            {
                testResult4.Text = DUT.PFPerTestItem[4] = DUT.FRTestStatus;

                if ((DUT.productModel == "100X-NAVY") || (productButton8.Checked)) //1000VB's
                {
                    if ((testResult3.Text == "FAILED") || (testResult4.Text == "FAILED")) reportPassFail3.Text = "FAILED";
                    else if ((testResult3.Text == "PASSED") && (testResult4.Text == "PASSED")) reportPassFail3.Text = "PASSED";
                    else reportPassFail3.Text = "NC";
                }
                else if (productButton20.Checked)
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        reportPassFail4.Text = testResult4.Text;
                    }
                    else reportPassFail4.Text = "NC";
                }
                else
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        if ((testResult3.Text == "FAILED") || (testResult4.Text == "FAILED") || (testResult5.Text == "FAILED")) reportPassFail3.Text = "FAILED";
                        else if ((testResult3.Text == "PASSED") && (testResult4.Text == "PASSED") && (testResult5.Text == "PASSED")) reportPassFail3.Text = "PASSED";
                    }
                    else reportPassFail3.Text = "NC";
                }

                if (DUT.FRTestStatus == "FAILED")
                {
                    testResult4.ForeColor = Color.Red;
                }
                else
                {
                    testResult4.ForeColor = Color.LightGreen;
                }
            }
            else if (testSelButton5.Checked)
            {
                testResult5.Text = DUT.PFPerTestItem[5] = DUT.FRTestStatus;

                if (DUT.productModel == "100X-NAVY")
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        reportPassFail4.Text = DUT.FRTestStatus;
                    }
                }
                else if (productButton20.Checked)
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        reportPassFail5.Text = testResult5.Text;
                    }
                    else reportPassFail5.Text = "NC";
                }
                else
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        if ((testResult3.Text == "FAILED") || (testResult4.Text == "FAILED") || (testResult5.Text == "FAILED")) reportPassFail3.Text = "FAILED";
                        else if ((testResult3.Text == "PASSED") && (testResult4.Text == "PASSED") && (testResult5.Text == "PASSED")) reportPassFail3.Text = "PASSED";
                    }
                    else reportPassFail3.Text = "NC";
                }

                if (DUT.FRTestStatus == "FAILED")
                {
                    testResult5.ForeColor = Color.Red;
                }
                else
                {
                    testResult5.ForeColor = Color.LightGreen;
                }
            }
            else if (testSelButton6.Checked)
            {
                testResult6.Text = DUT.PFPerTestItem[6] = DUT.NoiseTestStatus;

                if ((DUT.productModel == "100X") || (DUT.productModel == "100X-NAVY-V1") || (DUT.productModel == "1000X2U") || (productButton18.Checked)) //including RXs
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        reportPassFail3.Text = DUT.NoiseTestStatus;
                    }
                    //    reportPassFail3.Text = "PASSED"; //SPDebug
                }
                else if (DUT.productModel == "100X-NAVY")
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        reportPassFail5.Text = DUT.NoiseTestStatus;
                    }
                }
                else
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        reportPassFail4.Text = DUT.NoiseTestStatus;
                    }
                }

                if (DUT.NoiseTestStatus == "FAILED")
                {
                    testResult6.ForeColor = Color.Red;
                }
                else
                {
                    testResult6.ForeColor = Color.LightGreen;
                }
            }
            else if (testSelButton7.Checked)
            {
                if (DUT.productModel == "100X-NAVY") //100X-NAVY IPOD Sen & Max SPL
                {
                    if ((DUT.FRTestStatus == "FAILED") || (DUT.SPLTestStatus == "FAILED"))
                    {
                        testResult7.ForeColor = Color.Red;
                        testResult7.Text = DUT.PFPerTestItem[7] = "FAILED";
                        if (!troubleShootingCheckBox.Checked)
                        {
                            reportPassFail6.Text = "FAILED";
                        }
                    }
                    if ((DUT.FRTestStatus == "PASSED") && (DUT.SPLTestStatus == "PASSED"))
                    {
                        testResult7.ForeColor = Color.LightGreen;
                        testResult7.Text = DUT.PFPerTestItem[7] = "PASSED";
                        if (!troubleShootingCheckBox.Checked)
                        {
                            reportPassFail6.Text = "PASSED";
                        }
                    }
                }
                else
                {
                    if (DUT.FRTestStatus == "FAILED")
                    {
                        testResult7.ForeColor = Color.Red;
                        testResult7.Text = DUT.PFPerTestItem[7] = "FAILED";
                    }
                    else
                    {
                        testResult7.ForeColor = Color.LightGreen;
                        testResult7.Text = DUT.PFPerTestItem[7] = "PASSED";
                    }

                    //For report data for 1000V --> AUX cable & Limited key SW
                    if (AUXCableOption.Checked)
                    {
                        if (!troubleShootingCheckBox.Checked)
                        {
                            if (testResult7.Text == "FAILED")
                            {
                                reportPassFail5.Text = "FAILED";
                            }
                            else if (testResult7.Text == "PASSED")
                            {
                                reportPassFail5.Text = "PASSED";
                            }
                        }
                        else
                        {
                            reportPassFail5.Text = "NC";
                        }
                    }
                    else
                    {
                        if (!troubleShootingCheckBox.Checked)
                        {
                            if ((testResult7.Text == "FAILED") || (testResult9.Text == "FAILED"))
                            {
                                reportPassFail5.Text = "FAILED";
                            }
                            else if ((testResult7.Text == "PASSED") && (testResult9.Text == "PASSED"))
                            {
                                reportPassFail5.Text = "PASSED";
                            }
                        }
                        else
                        {
                            reportPassFail5.Text = "NC";
                        }
                    }
                }
            }
            else if (testSelButton8.Checked)
            {
                if (!troubleShootingCheckBox.Checked)
                {
                    reportPassFail6.Text = DUT.PFPerTestItem[8] = testResult8.Text;
                }
            }
            else if (testSelButton9.Checked)
            {
                if (productButton18.Checked) //if RX's
                {
                    if (!troubleShootingCheckBox.Checked)
                    {
                        reportPassFail2.Text = DUT.PFPerTestItem[9] = testResult9.Text;
                    }
                    //    reportPassFail2.Text = "PASSED"; //SPDebug
                    return;
                }

                if (AUXCableOption.Checked)
                {
                    DUT.FRTestStatus2 = "PASSED"; //Do not evaluate AUX cable
                }

                if ((DUT.FRTestStatus == "PASSED") && (DUT.FRTestStatus2 == "PASSED"))
                {
                    testResult9.ForeColor = Color.LightGreen;
                    testResult9.Text = DUT.PFPerTestItem[9] = "PASSED";
                }
                else if ((DUT.FRTestStatus == "FAILED") || (DUT.FRTestStatus2 == "FAILED"))
                {
                    testResult9.ForeColor = Color.Red;
                    testResult9.Text = DUT.PFPerTestItem[9] = "FAILED";
                }

                if (!troubleShootingCheckBox.Checked)
                {
                    if ((testResult7.Text == "FAILED") || (testResult9.Text == "FAILED"))
                    {
                        reportPassFail5.Text = "FAILED";
                    }
                    else if ((testResult7.Text == "PASSED") && (testResult9.Text == "PASSED"))
                    {
                        reportPassFail5.Text = "PASSED";
                    }
                }
                else if (troubleShootingCheckBox.Checked)
                {

                }
                else
                {
                    reportPassFail5.Text = "NC";
                }
            }

        }

        private bool DuplicatedSNWasFound()
        {
            Boolean SNFound = false;

            var SNcolumn = new List<string>();
            string LogFileName;

            if ((DUT.SN[14] == "productButton12") || (DUT.SN[14] == "productButton13")) LogFileName = "60-DEG";
            else if (DUT.SN[14] == "productButton5") LogFileName = "1000X";
            else if (DUT.SN[14] == "productButton11") LogFileName = "60-DEG w AmpPack";
            else if (DUT.SN[14] == "productButton14") LogFileName = "60-DEG wo Amp";
            else if (DUT.SN[14] == "productButton18") LogFileName = "LRAD-RX";
            else if (DUT.SN[14] == "productButton8") LogFileName = "1000XVB";
            else if (DUT.SN[14] == "productButton9") LogFileName = "1950XL";
            else if (DUT.SN[14] == "productButton10") LogFileName = "360X Manifold";
            else if (DUT.SN[14] == "productButton2") LogFileName = "450XL";
            else if (DUT.SN[14] == "productButton12") LogFileName = "60-DEG";
            else if ((DUT.SN[14] == "productButton20") || (DUT.SN[14] == "productButton21")) LogFileName = "360Xm";
            else if (DUT.SN[14] == "productButton21") LogFileName = "360Xm wo Trans";
            else if (DUT.SN[14] == "productButton22") LogFileName = "360XL-MID";
            else LogFileName = DUT.productModel;

            try
            {
                using (var rd = new StreamReader(Folder.networkDrive + LogFileName + " Test Log (Post-Test).csv"))
                {
                    while (!rd.EndOfStream)
                    {
                        var splits = rd.ReadLine().Split(',');
                        SNcolumn.Add(splits[0]);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true; //This might trigger another popup warning duplicate SN found              
            }

            foreach (var element in SNcolumn)
            {
                if (element == DUT.SN[1])
                {
                    SNFound = true;
                    break;
                }
            }

            if ((SNFound) && (remarkBox.Text == ""))
            {
                DialogResult dialogResult = MessageBox.Show("The sample SN was detected in the serial number system,\n " +
                  "Please enter a reason for a retest.", "Message",
                  MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else SNFound = false;

            return SNFound;
        }

        private string GetRefFileName()
        {
            string RefFileName = null, ModelName = null;

            if (DUT.productModel == "450XL Extended Test") ModelName = "450XL";
            else ModelName = DUT.productModel;

            if (testSelButton0.Checked)
            {
                RefFileName = ModelName + " Test 0 Ref.xlsx";
            }
            else if (testSelButton1.Checked)
            {
                RefFileName = ModelName + " Test 1 Ref.xlsx";
            }
            else if (testSelButton2.Checked)
            {
                RefFileName = ModelName + " Test 2 Ref.xlsx";
            }
            else if (testSelButton3.Checked)
            {
                RefFileName = ModelName + " Test 3 Ref.xlsx";
            }
            else if (testSelButton4.Checked)
            {
                RefFileName = ModelName + " Test 4 Ref.xlsx";
            }
            else if (testSelButton5.Checked)
            {
                RefFileName = ModelName + " Test 5 Ref.xlsx";
            }
            else if (testSelButton6.Checked)
            {
                RefFileName = ModelName + " Test 6 Ref.xlsx";
            }
            else if (testSelButton7.Checked)
            {
                RefFileName = ModelName + " Test 7 Ref.xlsx";
            }
            else if (testSelButton9.Checked)
            {
                RefFileName = ModelName + " Test 9 Ref.xlsx";
            }

            return RefFileName;
        }

        private string GetRefFileNamePlus()
        {
            string RefFileName = null, ModelName = null;

            if (DUT.productModel == "450XL Extended Test") ModelName = "450XL";
            else ModelName = DUT.productModel;

            if (testSelButton0.Checked)
            {
                RefFileName = ModelName + " Test 0 Ref+.xlsx";
            }
            else if (testSelButton1.Checked)
            {
                RefFileName = ModelName + " Test 1 Ref+.xlsx";
            }
            else if (testSelButton2.Checked)
            {
                RefFileName = ModelName + " Test 2 Ref+.xlsx";
            }
            else if (testSelButton3.Checked)
            {
                RefFileName = ModelName + " Test 3 Ref+.xlsx";
            }
            else if (testSelButton4.Checked)
            {
                RefFileName = ModelName + " Test 4 Ref+.xlsx";
            }
            else if (testSelButton5.Checked)
            {
                RefFileName = ModelName + " Test 5 Ref+.xlsx";
            }
            else if (testSelButton6.Checked)
            {
                RefFileName = ModelName + " Test 6 Ref+.xlsx";
            }
            else if (testSelButton7.Checked)
            {
                RefFileName = ModelName + " Test 7 Ref+.xlsx";
            }
            else if (testSelButton9.Checked)
            {
                RefFileName = ModelName + " Test 9 Ref+.xlsx";
            }

            return RefFileName;
        }

        private string GetRefFileNameMinus()
        {
            string RefFileName = null, ModelName = null;

            if (DUT.productModel == "450XL Extended Test") ModelName = "450XL";
            else ModelName = DUT.productModel;

            if (testSelButton0.Checked)
            {
                RefFileName = ModelName + " Test 0 Ref-.xlsx";
            }
            else if (testSelButton1.Checked)
            {
                RefFileName = ModelName + " Test 1 Ref-.xlsx";
            }
            else if (testSelButton2.Checked)
            {
                RefFileName = ModelName + " Test 2 Ref-.xlsx";
            }
            else if (testSelButton3.Checked)
            {
                RefFileName = ModelName + " Test 3 Ref-.xlsx";
            }
            else if (testSelButton4.Checked)
            {
                RefFileName = ModelName + " Test 4 Ref-.xlsx";
            }
            else if (testSelButton5.Checked)
            {
                RefFileName = ModelName + " Test 5 Ref-.xlsx";
            }
            else if (testSelButton6.Checked)
            {
                RefFileName = ModelName + " Test 6 Ref-.xlsx";
            }
            else if (testSelButton7.Checked)
            {
                RefFileName = ModelName + " Test 7 Ref-.xlsx";
            }
            else if (testSelButton9.Checked)
            {
                RefFileName = ModelName + " Test 9 Ref-.xlsx";
            }

            return RefFileName;
        }

        void GetTestParameters(String DriveLevelFR, String DriveLevelSPL, String RefLower, String RefUpper, String WindowingBegin, String WindowingEnd, String FreqSPL)
        {
            string ModelName = null;

            if (DUT.productModel == "450XL Extended Test") ModelName = "450XL";
            else ModelName = DUT.productModel;

            try
            {
                DUT.stimulusFR = Convert.ToDouble(ini.IniReadValue(ModelName, DriveLevelFR));
            }
            catch (Exception)
            {
                DUT.stimulusFR = Convert.ToDouble(ini.IniReadValue("RefData", DriveLevelFR));
            }

            try
            {
                DUT.stimulusSPL = Convert.ToDouble(ini.IniReadValue(ModelName, DriveLevelSPL));
            }
            catch (Exception)
            {
                DUT.stimulusSPL = Convert.ToDouble(ini.IniReadValue("RefData", DriveLevelSPL));
            }

            try
            {
                DUT.SweepLowerLimit = Convert.ToDouble(ini.IniReadValue(ModelName, RefLower));
            }
            catch (Exception)
            {
                DUT.SweepLowerLimit = Convert.ToDouble(ini.IniReadValue("RefData", RefLower));
            }

            try
            {
                DUT.SweepUpperLimit = Convert.ToDouble(ini.IniReadValue(ModelName, RefUpper));
            }
            catch (Exception)
            {
                DUT.SweepUpperLimit = Convert.ToDouble(ini.IniReadValue("RefData", RefUpper));
            }

            try
            {
                DUT.WindowBegin = Convert.ToDouble(ini.IniReadValue(ModelName, WindowingBegin));
            }
            catch (Exception)
            {
                DUT.WindowBegin = Convert.ToDouble(ini.IniReadValue("RefData", WindowingBegin));
            }

            try
            {
                DUT.WindowEnd = Convert.ToDouble(ini.IniReadValue(ModelName, WindowingEnd));
            }
            catch (Exception)
            {
                DUT.WindowEnd = Convert.ToDouble(ini.IniReadValue("RefData", WindowingEnd));
            }

            try
            {
                DUT.freqStart = Convert.ToInt16(ini.IniReadValue(ModelName, "FreqStart"));
            } //For FR sweep range
            catch (Exception)
            {
                DUT.freqStart = Convert.ToInt16(ini.IniReadValue("RefData", "FreqStart"));
            }

            try
            {
                DUT.freqStop = Convert.ToInt16(ini.IniReadValue(ModelName, "FreqStop"));
            } //For FR sweep range
            catch (Exception)
            {
                DUT.freqStop = Convert.ToInt16(ini.IniReadValue("RefData", "FreqStop"));
            }

            try
            {
                DUT.freqSPL = Convert.ToInt16(ini.IniReadValue(ModelName, FreqSPL));
            }
            catch (Exception)
            {
                DUT.freqSPL = Convert.ToInt16(ini.IniReadValue("RefData", FreqSPL));
            }

            if (testSelButton0.Checked)
            {
                try
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue(ModelName, "RefSPL"));
                }
                catch (Exception)
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue("RefData", "RefSPL"));
                }

                try
                {
                    DUT.SPLLowerLimit = Convert.ToDouble(ini.IniReadValue(ModelName, "SPLLowerLimit"));
                }
                catch (Exception)
                {
                    DUT.SPLLowerLimit = Convert.ToDouble(ini.IniReadValue("RefData", "SPLLowerLimit"));
                }
                try
                {
                    DUT.SPLUpperLimit = Convert.ToDouble(ini.IniReadValue(ModelName, "SPLUpperLimit"));
                }
                catch (Exception)
                {
                    DUT.SPLUpperLimit = Convert.ToDouble(ini.IniReadValue("RefData", "SPLUpperLimit"));
                }

                //   remarkBox.Text = "Ref=" + Convert.ToString(DUT.RefSPL) + ", Lower=" + Convert.ToString(DUT.SPLLowerLimit) + ", Upper=" + Convert.ToString(DUT.SPLUpperLimit);

                DUT.curveSmooth = ini.IniReadValue(ModelName, "curveSmooth");
                if (DUT.curveSmooth == "")
                {
                    DUT.curveSmooth = ini.IniReadValue("RefData", "curveSmooth");
                }

                DUT.graphTitle = ini.IniReadValue(ModelName, "Test0_FR_graphTitle");
                if (DUT.graphTitle == "")
                {
                    DUT.graphTitle = ini.IniReadValue("RefData", "Test0_FR_graphTitle");
                }

                DUT.SPLGraphTitle = ini.IniReadValue(ModelName, "Test0_SPL_graphTitle");
                if (DUT.SPLGraphTitle == "")
                {
                    DUT.SPLGraphTitle = ini.IniReadValue("RefData", "Test0_SPL_graphTitle");
                }

                try
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMin"));
                }
                catch (Exception)
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMin"));
                }
                try
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMax"));
                }
                catch (Exception)
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMax"));
                }

                try
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMin"));
                }
                catch (Exception)
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMin"));
                }
                try
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMax"));
                }
                catch (Exception)
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMax"));
                }

                try
                {
                    DUT.YAxisMaxSPL = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMaxSPL"));
                }
                catch (Exception)
                {
                    DUT.YAxisMaxSPL = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMaxSPL"));
                }
                try
                {
                    DUT.YAxisMinSPL = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMinSPL"));
                }
                catch (Exception)
                {
                    DUT.YAxisMinSPL = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMinSPL"));
                }
            }
            else if (testSelButton1.Checked)
            {

                try
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue(ModelName, "RefSPL1"));
                }
                catch (Exception)
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue("RefData", "RefSPL1"));
                }

                try
                {
                    DUT.SPLLowerLimit = Convert.ToDouble(ini.IniReadValue(ModelName, "SPLLowerLimit1"));
                } //Only for SS400 accelerometer output
                catch (Exception)
                {
                    DUT.SPLLowerLimit = Convert.ToDouble(ini.IniReadValue("RefData", "SPLLowerLimit1"));
                }
                try
                {
                    DUT.SPLUpperLimit = Convert.ToDouble(ini.IniReadValue(ModelName, "SPLUpperLimit1"));
                }
                catch (Exception)
                {
                    DUT.SPLUpperLimit = Convert.ToDouble(ini.IniReadValue("RefData", "SPLUpperLimit1"));
                }

                DUT.curveSmooth = ini.IniReadValue(ModelName, "curveSmooth1");
                if (DUT.curveSmooth == "")
                {
                    DUT.curveSmooth = ini.IniReadValue("RefData", "curveSmooth1");
                }

                DUT.graphTitle = ini.IniReadValue(ModelName, "Test1_FR_graphTitle");
                if (DUT.graphTitle == "")
                {
                    DUT.graphTitle = ini.IniReadValue("RefData", "Test1_FR_graphTitle");
                }

                DUT.SPLGraphTitle = ini.IniReadValue(ModelName, "Test1_SPL_graphTitle");
                if (DUT.SPLGraphTitle == "")
                {
                    DUT.SPLGraphTitle = ini.IniReadValue("RefData", "Test1_SPL_graphTitle");
                }

                try
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMin1"));
                }
                catch (Exception)
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMin1"));
                }
                try
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMax1"));
                }
                catch (Exception)
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMax1"));
                }

                try
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMin1"));
                }
                catch (Exception)
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMin1"));
                }
                try
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMax1"));
                }
                catch (Exception)
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMax1"));
                }

                try
                {
                    DUT.YAxisMaxSPL = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMaxSPL1"));
                } //Graphics scal for RX only
                catch (Exception)
                {
                    DUT.YAxisMaxSPL = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMaxSPL1"));
                }
                try
                {
                    DUT.YAxisMinSPL = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMinSPL1"));
                }
                catch (Exception)
                {
                    DUT.YAxisMinSPL = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMinSPL1"));
                }
            }
            else if (testSelButton2.Checked)
            {
                try
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue(ModelName, "RefSPL2"));
                }
                catch (Exception)
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue("RefData", "RefSPL2"));
                }

                try
                {
                    DUT.SPLLowerLimit = Convert.ToDouble(ini.IniReadValue(ModelName, "SPLLowerLimit2"));
                } //Only for RX Alert Tone SPL
                catch (Exception)
                {
                    DUT.SPLLowerLimit = Convert.ToDouble(ini.IniReadValue("RefData", "SPLLowerLimit2"));
                }
                try
                {
                    DUT.SPLUpperLimit = Convert.ToDouble(ini.IniReadValue(ModelName, "SPLUpperLimit2"));
                }
                catch (Exception)
                {
                    DUT.SPLUpperLimit = Convert.ToDouble(ini.IniReadValue("RefData", "SPLUpperLimit2"));
                }

                DUT.curveSmooth = ini.IniReadValue(ModelName, "curveSmooth2");
                if (DUT.curveSmooth == "")
                {
                    DUT.curveSmooth = ini.IniReadValue("RefData", "curveSmooth2");
                }

                DUT.graphTitle = ini.IniReadValue(ModelName, "Test2_FR_graphTitle");
                if (DUT.graphTitle == "")
                {
                    DUT.graphTitle = ini.IniReadValue("RefData", "Test2_FR_graphTitle");
                }

                DUT.SPLGraphTitle = ini.IniReadValue(ModelName, "Test2_SPL_graphTitle");
                if (DUT.SPLGraphTitle == "")
                {
                    DUT.SPLGraphTitle = ini.IniReadValue("RefData", "Test2_SPL_graphTitle");
                }

                try
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMin2"));
                }
                catch (Exception)
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMin2"));
                }
                try
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMax2"));
                }
                catch (Exception)
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMax2"));
                }

                try
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMin2"));
                }
                catch (Exception)
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMin2"));
                }
                try
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMax2"));
                }
                catch (Exception)
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMax2"));
                }
            }
            else if (testSelButton3.Checked)
            {
                try
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue(ModelName, "RefSPL3"));
                }
                catch (Exception)
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue("RefData", "RefSPL3"));
                }

                try
                {
                    DUT.SPLLowerLimit = Convert.ToDouble(ini.IniReadValue(ModelName, "SPLLowerLimit3"));
                }
                catch (Exception)
                {
                    DUT.SPLLowerLimit = Convert.ToDouble(ini.IniReadValue("RefData", "SPLLowerLimit3"));
                }
                try
                {
                    DUT.SPLUpperLimit = Convert.ToDouble(ini.IniReadValue(ModelName, "SPLUpperLimit3"));
                }
                catch (Exception)
                {
                    DUT.SPLUpperLimit = Convert.ToDouble(ini.IniReadValue("RefData", "SPLUpperLimit3"));
                }

                DUT.curveSmooth = ini.IniReadValue(ModelName, "curveSmooth3");
                if (DUT.curveSmooth == "")
                {
                    DUT.curveSmooth = ini.IniReadValue("RefData", "curveSmooth3");
                }

                DUT.graphTitle = ini.IniReadValue(ModelName, "Test3_FR_graphTitle");
                if (DUT.graphTitle == "")
                {
                    DUT.graphTitle = ini.IniReadValue("RefData", "Test3_FR_graphTitle");
                }

                DUT.SPLGraphTitle = ini.IniReadValue(ModelName, "Test3_SPL_graphTitle");
                if (DUT.SPLGraphTitle == "")
                {
                    DUT.SPLGraphTitle = ini.IniReadValue("RefData", "Test3_SPL_graphTitle");
                }

                try
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMin3"));
                }
                catch (Exception)
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMin3"));
                }
                try
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMax3"));
                }
                catch (Exception)
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMax3"));
                }

                try
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMin3"));
                }
                catch (Exception)
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMin3"));
                }
                try
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMax3"));
                }
                catch (Exception)
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMax3"));
                }
            }
            else if (testSelButton4.Checked)
            {
                DUT.curveSmooth = ini.IniReadValue(ModelName, "curveSmooth4");
                if (DUT.curveSmooth == "")
                {
                    DUT.curveSmooth = ini.IniReadValue("RefData", "curveSmooth4");
                }

                DUT.graphTitle = ini.IniReadValue(ModelName, "Test4_FR_graphTitle");
                if (DUT.graphTitle == "")
                {
                    DUT.graphTitle = ini.IniReadValue("RefData", "Test4_FR_graphTitle");
                }

                try
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMin4"));
                }
                catch (Exception)
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMin4"));
                }
                try
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMax4"));
                }
                catch (Exception)
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMax4"));
                }

                try
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMin4"));
                }
                catch (Exception)
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMin4"));
                }
                try
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMax4"));
                }
                catch (Exception)
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMax4"));
                }
            }
            else if (testSelButton5.Checked)
            {
                DUT.curveSmooth = ini.IniReadValue(ModelName, "curveSmooth5");
                if (DUT.curveSmooth == "")
                {
                    DUT.curveSmooth = ini.IniReadValue("RefData", "curveSmooth5");
                }

                DUT.graphTitle = ini.IniReadValue(ModelName, "Test5_FR_graphTitle");
                if (DUT.graphTitle == "")
                {
                    DUT.graphTitle = ini.IniReadValue("RefData", "Test5_FR_graphTitle");
                }

                try
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMin5"));
                }
                catch (Exception)
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMin5"));
                }
                try
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMax5"));
                }
                catch (Exception)
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMax5"));
                }

                try
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMin5"));
                }
                catch (Exception)
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMin5"));
                }
                try
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMax5"));
                }
                catch (Exception)
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMax5"));
                }
            }
            else if (testSelButton6.Checked)
            {
                //       try { DUT.RefSPL = Convert.ToDouble(ini.IniReadValue(ModelName, "RefNoise")); }
                //       catch (Exception) { DUT.RefSPL = Convert.ToDouble(ini.IniReadValue("RefData", "RefNoise")); }   //NOT need, noise test was based on noise profile across Freq Spectrum

                DUT.graphTitle = ini.IniReadValue(ModelName, "Test6_FR_graphTitle");
                if (DUT.graphTitle == "")
                {
                    DUT.graphTitle = ini.IniReadValue("RefData", "Test6_FR_graphTitle");
                }

                try
                {
                    DUT.SPLUpperLimit = Convert.ToDouble(ini.IniReadValue(ModelName, "SPLUpperLimit6"));
                }
                catch (Exception)
                {
                    DUT.SPLUpperLimit = Convert.ToDouble(ini.IniReadValue("RefData", "SPLUpperLimit6"));
                }
                try
                {
                    DUT.SPLLowerLimit = Convert.ToDouble(ini.IniReadValue(ModelName, "SPLLowerLimit6"));
                }
                catch (Exception)
                {
                    DUT.SPLLowerLimit = Convert.ToDouble(ini.IniReadValue("RefData", "SPLLowerLimit6"));
                }

                try
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMin6"));
                }
                catch (Exception)
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMin6"));
                }
                try
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMax6"));
                }
                catch (Exception)
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMax6"));
                }

                try
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMin6"));
                }
                catch (Exception)
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMin6"));
                }
                try
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMax6"));
                }
                catch (Exception)
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMax6"));
                }
            }
            else if (testSelButton7.Checked)
            {
                try
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue(ModelName, "RefSPL7"));
                }
                catch (Exception)
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue("RefData", "RefSPL7"));
                }

                DUT.curveSmooth = ini.IniReadValue(ModelName, "curveSmooth7");
                if (DUT.curveSmooth == "")
                {
                    DUT.curveSmooth = ini.IniReadValue("RefData", "curveSmooth7");
                }

                DUT.graphTitle = ini.IniReadValue(ModelName, "Test7_FR_graphTitle");
                if (DUT.graphTitle == "")
                {
                    DUT.graphTitle = ini.IniReadValue("RefData", "Test7_FR_graphTitle");
                }

                try
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMin7"));
                }
                catch (Exception)
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMin7"));
                }
                try
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMax7"));
                }
                catch (Exception)
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMax7"));
                }

                try
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMin7"));
                }
                catch (Exception)
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMin7"));
                }
                try
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMax7"));
                }
                catch (Exception)
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMax7"));
                }
            }
            else if (testSelButton8.Checked) { }
            else if (testSelButton9.Checked)
            {
                DUT.curveSmooth = ini.IniReadValue(ModelName, "curveSmooth9");
                if (DUT.curveSmooth == "")
                {
                    DUT.curveSmooth = ini.IniReadValue("RefData", "curveSmooth9");
                }

                DUT.graphTitle = ini.IniReadValue(ModelName, "Test9_FR_graphTitle");
                if (DUT.graphTitle == "")
                {
                    DUT.graphTitle = ini.IniReadValue("RefData", "Test9_FR_graphTitle");
                }

                try
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMin9"));
                }
                catch (Exception)
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMin9"));
                }
                try
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "XAxisMax9"));
                }
                catch (Exception)
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMax9"));
                }

                try
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMin9"));
                }
                catch (Exception)
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMin9"));
                }
                try
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue(ModelName, "YAxisMax9"));
                }
                catch (Exception)
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMax9"));
                }
            }
        }

        private void RetriveTestSetup(object sender, EventArgs e)
        {
            try
            {
                DUT.freqStart = Convert.ToInt16(ini.IniReadValue(DUT.productModel, "FreqStart"));
            } //For FR sweep range
            catch (Exception)
            {
                DUT.freqStart = Convert.ToInt16(ini.IniReadValue("RefData", "FreqStart"));
            }

            try
            {
                DUT.freqStop = Convert.ToInt16(ini.IniReadValue(DUT.productModel, "FreqStop"));
            } //For FR sweep range
            catch (Exception)
            {
                DUT.freqStop = Convert.ToInt16(ini.IniReadValue("RefData", "FreqStop"));
            }

            if (testSelButton0.Checked)
            {
                setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction0");
                if (setupInstruction.Text == "")
                {
                    setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction0");
                }

                DUT.curveSmooth = ini.IniReadValue(DUT.productModel, "curveSmooth");
                if (DUT.curveSmooth == "")
                {
                    DUT.curveSmooth = ini.IniReadValue("RefData", "curveSmooth");
                }

                try
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue(DUT.productModel, "Ref_SPL"));
                }
                catch (Exception)
                {
                    DUT.RefSPL = Convert.ToDouble(ini.IniReadValue("RefData", "Ref_SPL"));
                }

                try
                {
                    DUT.RefSS400MicInput = Convert.ToDouble(ini.IniReadValue(DUT.productModel, "Ref_SS400MicInput"));
                }
                catch (Exception)
                {
                    DUT.RefSS400MicInput = Convert.ToDouble(ini.IniReadValue("RefData", "Ref_SS400MicInput"));
                }

                DUT.graphTitle = ini.IniReadValue(DUT.productModel, "Test0_FR_graphTitle");
                if (DUT.graphTitle == "")
                {
                    DUT.graphTitle = ini.IniReadValue("RefData", "Test0_FR_graphTitle");
                }
                DUT.SPLGraphTitle = ini.IniReadValue(DUT.productModel, "Test0_SPL_graphTitle");
                if (DUT.SPLGraphTitle == "")
                {
                    DUT.SPLGraphTitle = ini.IniReadValue("RefData", "Test0_SPL_graphTitle");
                }
                DUT.SS400MicInputGraphTitle = ini.IniReadValue(DUT.productModel, "Test0_MIC_graphTitle");
                if (DUT.SS400MicInputGraphTitle == "")
                {
                    DUT.SS400MicInputGraphTitle = ini.IniReadValue("RefData", "Test0_MIC_graphTitle");
                }

                try
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue(DUT.productModel, "XAxisMin"));
                }
                catch (Exception)
                {
                    DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMin"));
                }
                try
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue(DUT.productModel, "XAxisMax"));
                }
                catch (Exception)
                {
                    DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "XAxisMax"));
                }

                try
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue(DUT.productModel, "YAxisMin"));
                }
                catch (Exception)
                {
                    DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMin"));
                }
                try
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue(DUT.productModel, "YAxisMax"));
                }
                catch (Exception)
                {
                    DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMax"));
                }

                try
                {
                    DUT.YAxisMaxSPL = Convert.ToDouble(ini.IniReadValue(DUT.productModel, "YAxisMaxSPL"));
                }
                catch (Exception)
                {
                    DUT.YAxisMaxSPL = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMaxSPL"));
                }
                try
                {
                    DUT.YAxisMinSPL = Convert.ToDouble(ini.IniReadValue(DUT.productModel, "YAxisMinSPL"));
                }
                catch (Exception)
                {
                    DUT.YAxisMinSPL = Convert.ToDouble(ini.IniReadValue("RefData", "YAxisMinSPL"));
                }

            }
            else
            {
                string Title = null, SPLTitle = null, YMin = null, YMax = null, XMin = null, XMax = null, CurveSmooth = null;

                if (testSelButton1.Checked)
                {
                    setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction1");
                    if (setupInstruction.Text == "")
                    {
                        setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction1");
                    }

                    Title = "Test1_graphTitle";
                    YMin = "YAxisMin1";
                    YMax = "YAxisMax1";
                    XMin = "XAxisMin1";
                    XMax = "XAxisMax1";
                    CurveSmooth = "curveSmooth1";
                }
                else if (testSelButton2.Checked)
                {
                    setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction2");
                    if (setupInstruction.Text == "")
                    {
                        setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction2");
                    }

                    Title = "Test2_graphTitle";
                    YMin = "YAxisMin2";
                    YMax = "YAxisMax2";
                    XMin = "XAxisMin2";
                    XMax = "XAxisMax2";
                    CurveSmooth = "curveSmooth2";
                }
                else if (testSelButton3.Checked)
                {
                    setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction3");
                    if (setupInstruction.Text == "")
                    {
                        setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction3");
                    }

                    Title = "Test3_graphTitle";
                    YMin = "YAxisMin3";
                    YMax = "YAxisMax3";
                    XMin = "XAxisMin3";
                    XMax = "XAxisMax3";
                    CurveSmooth = "curveSmooth3";
                }
                else if (testSelButton4.Checked)
                {
                    setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction4");
                    if (setupInstruction.Text == "")
                    {
                        setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction4");
                    }

                    Title = "Test4_graphTitle";
                    YMin = "YAxisMin4";
                    YMax = "YAxisMax4";
                    XMin = "XAxisMin4";
                    XMax = "XAxisMax4";
                    CurveSmooth = "curveSmooth4";
                }
                else if (testSelButton5.Checked)
                {
                    setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction5");
                    if (setupInstruction.Text == "")
                    {
                        setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction5");
                    }

                    Title = "Test5_graphTitle";
                    YMin = "YAxisMin5";
                    YMax = "YAxisMax5";
                    XMin = "XAxisMin5";
                    XMax = "XAxisMax5";
                    CurveSmooth = "curveSmooth5";
                }
                else if (testSelButton6.Checked)
                {
                    setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction6");
                    if (setupInstruction.Text == "")
                    {
                        setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction6");
                    }

                    Title = "Test6_graphTitle";
                    YMin = "YAxisMin6";
                    YMax = "YAxisMax6";
                    XMin = "XAxisMin6";
                    XMax = "XAxisMax6";
                    CurveSmooth = "curveSmooth6";
                }
                else if (testSelButton7.Checked)
                {
                    setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction7");
                    if (setupInstruction.Text == "")
                    {
                        setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction7");
                    }

                    Title = "Test7_graphTitle";
                    SPLTitle = "Test7_SPL_graphTitle";
                    YMin = "YAxisMin7";
                    YMax = "YAxisMax7";
                    XMin = "XAxisMin7";
                    XMax = "XAxisMax7";
                    CurveSmooth = "curveSmooth7";
                }
                else if (testSelButton8.Checked)
                {
                    setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction8");
                    if (setupInstruction.Text == "")
                    {
                        setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction8");
                    }
                }
                else if (testSelButton9.Checked)
                {
                    setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction9");
                    if (setupInstruction.Text == "")
                    {
                        setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction9");
                    }

                    Title = "Test9_graphTitle";
                    YMin = "YAxisMin9";
                    YMax = "YAxisMax9";
                    XMin = "XAxisMin9";
                    XMax = "XAxisMax9";
                    CurveSmooth = "curveSmooth9";
                }

                if (!testSelButton8.Checked)
                {
                    DUT.graphTitle = ini.IniReadValue(DUT.productModel, Title);
                    if (DUT.graphTitle == "")
                    {
                        DUT.graphTitle = ini.IniReadValue("RefData", Title);
                    }

                    DUT.SPLGraphTitle = ini.IniReadValue(DUT.productModel, SPLTitle);
                    if (DUT.SPLGraphTitle == "")
                    {
                        DUT.SPLGraphTitle = ini.IniReadValue("RefData", SPLTitle);
                    }

                    try
                    {
                        DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue(DUT.productModel, YMin));
                    }
                    catch (Exception)
                    {
                        DUT.YAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", YMin));
                    }
                    try
                    {
                        DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue(DUT.productModel, YMax));
                    }
                    catch (Exception)
                    {
                        DUT.YAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", YMax));
                    }

                    try
                    {
                        DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue(DUT.productModel, XMin));
                    }
                    catch (Exception)
                    {
                        DUT.XAxisMin = Convert.ToDouble(ini.IniReadValue("RefData", XMin));
                    }
                    try
                    {
                        DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue(DUT.productModel, XMax));
                    }
                    catch (Exception)
                    {
                        DUT.XAxisMax = Convert.ToDouble(ini.IniReadValue("RefData", XMax));
                    }

                    DUT.curveSmooth = ini.IniReadValue(DUT.productModel, CurveSmooth);
                    if (DUT.curveSmooth == "")
                    {
                        DUT.curveSmooth = ini.IniReadValue("RefData", CurveSmooth);
                    }
                }

            }
        }

        private string ProductSeriesID()
        {
            if (productButton0.Checked) return "productButton0";
            if (productButton1.Checked) return "productButton1";
            if (productButton2.Checked) return "productButton2";
            if (productButton3.Checked) return "productButton3";
            if (productButton4.Checked) return "productButton4";
            if (productButton5.Checked) return "productButton5";
            if (productButton6.Checked) return "productButton6";
            if (productButton7.Checked) return "productButton7";
            if (productButton8.Checked) return "productButton8";
            if (productButton9.Checked) return "productButton9";
            if (productButton10.Checked) return "productButton10";
            if (productButton11.Checked) return "productButton11";
            if (productButton12.Checked) return "productButton12";
            if (productButton13.Checked) return "productButton13";
            if (productButton14.Checked) return "productButton14";
            if (productButton15.Checked) return "productButton15";
            if (productButton16.Checked) return "productButton16";
            if (productButton17.Checked) return "productButton17";
            if (productButton18.Checked) return "productButton18";
            if (productButton19.Checked) return "productButton19";
            if (productButton20.Checked) return "productButton20";
            if (productButton21.Checked) return "productButton21";
            if (productButton22.Checked) return "productButton22";
            else return "productButton0";
        }

        void recordSampleInfo(bool troubleShootingMode)
        {
            DUT.SN[0] = oper_Ini.Text;
            reportOperIni.Text = "Test Operator: " + oper_Ini.Text;

            if (troubleShootingMode)
            {
                DUT.SN[1] = "Screening";
            }
            else
            {
                DUT.SN[1] = SN_Box1.Text;
            }

            DUT.SN[2] = SN_Box2.Text;
            DUT.SN[3] = SN_Box3.Text;
            DUT.SN[4] = SN_Box4.Text;
            DUT.SN[5] = SN_Box5.Text;
            DUT.SN[6] = SN_Box6.Text;
            DUT.SN[7] = SN_Box7.Text;
            DUT.SN[8] = SN_Box8.Text;
            DUT.SN[9] = SN_Box9.Text;
            DUT.SN[10] = SN_Box10.Text;

            DUT.SN[14] = ProductSeriesID();
            DUT.SN[15] = DUT.productModel;
            DUT.WO_No = WO_No.Text;

            DUT.driverSN[1] = driverSN_Box1.Text;
            DUT.driverSN[2] = driverSN_Box2.Text;
            DUT.driverSN[3] = driverSN_Box3.Text;
            DUT.driverSN[4] = driverSN_Box4.Text;
            DUT.driverSN[5] = driverSN_Box5.Text;
            DUT.driverSN[6] = driverSN_Box6.Text;
            DUT.driverSN[7] = driverSN_Box7.Text;
            DUT.driverSN[8] = driverSN_Box8.Text;

            testTimeStamp.Text = date_Stamp.Text + "  " + time_Stamp.Text;
            //   testTimeStamp.Text = remarkBox.Text;  //SPDebug

            reportTextBox1.Text = reportTextBox2.Text = reportTextBox3.Text = reportTextBox4.Text = null;
            reportTextBox5.Text = reportTextBox6.Text = reportTextBox7.Text = null;

            if (preTest.Checked) DUT.FileNameExt = " (Pre-Test)";
            else if (postTest.Checked) DUT.FileNameExt = " (Post-Test)";
            else DUT.FileNameExt = " (Repaired)";

            label16.Text = "GENASYS PRODUCT TRAVELER";
            ReportFormPN.Text = DUT.ReportFormPN;

            if (troubleShootingMode)
            {
                reportModel.ForeColor = Color.Red;
                reportModel.Text = "Model: " + DUT.productModel + " - Trouble Shooting Mode";
                return;
            }

            reportModel.ForeColor = Color.Black;
            if (DUT.FileNameExt == " (Pre-Test)") reportModel.Text = "Model: " + DUT.productModel + DUT.FileNameExt;
            else reportModel.Text = "Model: " + DUT.productModel;

            //     if (DUT.productModel == "1000XVB-100FT") reportModel.Text = "Model: LRAD 1000X";  //Debug

            /*******************************************************************************************************************************************************************************/
            /* 1St sesion with S/Ns 
            /*******************************************************************************************************************************************************************************/
            if (DUT.productModel == "1000")
            {
                label16.Text = "ACCEPTANCE TEST REPORT";
                reportTextBox1.Text = "  LRAD-1000-G-SYS";
                reportTextBox2.Text = "  LRAD-1000-AHD-G";
                reportTextBox3.Text = "  LRAD-1000-AMP";
                reportTextBox4.Text = "  LRAD-1000-AC-PWR";
                reportTextBox5.Text = "  LRAD-X-MP3-AL";
                reportTextBox6.Text = "  LRAD-X-MIC-AL";
                reportTextBox7.Text = "  LRAD-PHRASLTR-2";

                reportSerialNo1.Text = SN_Box5.Text;
                reportSerialNo2.Text = SN_Box1.Text;
                reportSerialNo3.Text = SN_Box3.Text;
                reportSerialNo4.Text = SN_Box6.Text;
                reportSerialNo5.Text = SN_Box2.Text;
                reportSerialNo6.Text = SN_Box4.Text;
                reportSerialNo7.Text = SN_Box7.Text;
            }
            else if (productButton18.Checked) //LRAD-RX
            {
                reportTextBox1.Text = "  SYSTEM";
                reportTextBox2.Text = "  SPEAKER No 1";
                reportTextBox3.Text = "  SPEAKER No 2";
                reportTextBox4.Text = "  ELECTRONICS";
                reportTextBox5.Text = "  CAMERA";
                reportTextBox6.Text = "  LIGHT";
                reportTextBox7.Text = "  48VDC PWR SUPPLY";

                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = SN_Box4.Text;
                reportSerialNo3.Text = SN_Box5.Text;
                reportSerialNo4.Text = SN_Box3.Text;
                reportSerialNo5.Text = SN_Box9.Text;
                reportSerialNo6.Text = SN_Box8.Text;
                reportSerialNo7.Text = SN_Box10.Text;

                if (DUT.productModel == "1000RX") reportSerialNo3.Text = "NA";
            }
            else if ((productButton1.Checked) || (productButton2.Checked)) //300X's & 450XL's
            {
                reportTextBox1.Text = "  SYSTEM";
                reportTextBox2.Text = "  CONTROL UNIT";
                reportTextBox3.Text = "  ELECTRONICS";
                reportTextBox4.Text = "  MICROPHONE";
                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = SN_Box2.Text;
                reportSerialNo3.Text = SN_Box3.Text;
                reportSerialNo4.Text = SN_Box4.Text;
            }
            else if (productButton3.Checked) //500X's
            {
                if (DUT.productModel == "500XRE")
                {
                    reportTextBox2.Text = "  CONTROL UNIT";
                }
                else
                {
                    reportTextBox2.Text = "  MP3 PLAYER";
                }

                reportTextBox1.Text = "  SYSTEM";
                reportTextBox3.Text = "  ELECTRONICS";
                reportTextBox4.Text = "  MICROPHONE";
                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = SN_Box2.Text;
                reportSerialNo3.Text = SN_Box3.Text;
                reportSerialNo4.Text = SN_Box4.Text;
            }
            else if ((DUT.productModel == "100X") || (DUT.productModel == "100X-NAVY-V1") || (DUT.productModel == "1000Xi"))
            {
                reportTextBox1.Text = "  SYSTEM";
                reportTextBox3.Text = "  MICROPHONE";
                if (DUT.productModel == "1000Xi")
                {
                    reportTextBox2.Text = "  CONTROL UNIT";
                }
                else
                {
                    reportTextBox2.Text = "  MP3 PLAYER";
                    reportTextBox4.Text = "  BATTERY";
                    reportSerialNo4.Text = SN_Box5.Text;
                }

                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = SN_Box2.Text;
                reportSerialNo3.Text = SN_Box4.Text;
            }
            else if (DUT.productModel == "100X-NAVY")
            {
                reportTextBox1.Text = "  SYSTEM";
                reportTextBox2.Text = "  MICROPHONE";
                reportTextBox3.Text = "  BATTERY";
                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = SN_Box4.Text;
                reportSerialNo3.Text = SN_Box5.Text;
            }
            else if ((productButton5.Checked) || (productButton8.Checked) || (productButton9.Checked) || (productButton11.Checked)) //1000X/1000XVB/1950XL/DS60 /w amppack
            {
                reportTextBox1.Text = "  SYSTEM";
                reportTextBox4.Text = "  MICROPHONE";

                if (productButton11.Checked)
                {
                    reportTextBox2.Text = "  CONTROL UNIT";
                } //DS60 /w amppack
                else
                {
                    reportTextBox2.Text = "  MP3 PLAYER";
                }

                if (productButton9.Checked)
                {
                    reportTextBox3.Text = "  AMP PACK";
                    reportTextBox2.Text = "  CONTROL UNIT";
                }
                else
                {
                    reportTextBox3.Text = "  AMP PACK";
                }

                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = SN_Box2.Text;
                reportSerialNo3.Text = SN_Box3.Text;
                reportSerialNo4.Text = SN_Box4.Text;
            }
            else if (DUT.productModel == "1000X2U")
            {
                reportTextBox1.Text = "  SYSTEM";
                reportTextBox2.Text = "  2U CHASSIS";
                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = SN_Box3.Text;
            }
            else if ((productButton12.Checked) || (productButton13.Checked) || (productButton14.Checked)) //DS60-X/XL w Trans and horn only
            {
                if (productButton14.Checked) //Horn only
                {
                    reportTextBox1.Text = "  SAMPLE";
                    reportTextBox2.Text = "  DRIVER";
                    reportSerialNo1.Text = SN_Box1.Text;
                    reportSerialNo2.Text = driverSN_Box1.Text;
                }
                else
                {
                    reportTextBox1.Text = "  SAMPLE";
                    reportTextBox2.Text = "  TRANSFORMER";
                    reportTextBox3.Text = "  DRIVER";
                    reportSerialNo1.Text = SN_Box1.Text;
                    reportSerialNo2.Text = SN_Box2.Text;
                    reportSerialNo3.Text = driverSN_Box1.Text;
                }
            }
            else if (productButton15.Checked) //SS
            {
                reportTextBox1.Text = "  SYSTEM";
                reportTextBox2.Text = "  DRIVER PANEL";
                if (DUT.productModel != "SS400")
                {
                    reportTextBox3.Text = "  TRANSFORMER";
                    reportSerialNo3.Text = SN_Box3.Text;
                }

                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = SN_Box2.Text;
            }
            else if (productButton16.Checked) //SSX w Trans
            {
                reportTextBox1.Text = "  SYSTEM";
                reportTextBox2.Text = "  TRANSFORMER";
                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = SN_Box2.Text;
            }
            else if ((DUT.productModel == "360Xm 1-ST w Trans") || (DUT.productModel == "360Xm 1-ST wo Trans") ||
              (DUT.productModel == "360Xm 1-ST 25V-60W") || (DUT.productModel == "360Xm 1-ST 100V-60W")) //360Xm 1-ST
            {
                reportTextBox1.Text = "  SYSTEM";
                reportTextBox2.Text = "  DRIVER";
                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = driverSN_Box1.Text;
            }
            else if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 100V-120W") ||
              (DUT.productModel == "360Xm 2-ST 70V-60W") || (DUT.productModel == "360Xm 2-ST wo Trans")) //360Xm 2-STF
            {
                reportTextBox1.Text = "  SYSTEM";
                reportTextBox2.Text = "  DRIVER NO 1";
                reportTextBox3.Text = "  DRIVER NO 2";
                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = driverSN_Box1.Text;
                reportSerialNo3.Text = driverSN_Box2.Text;
            }
            else if (DUT.productModel == "360Xm 4-ST 100V-240W") //360Xm 4-ST                    
            {
                reportTextBox1.Text = "  SYSTEM";
                reportTextBox2.Text = "  DRIVER NO 1";
                reportTextBox3.Text = "  DRIVER NO 2";
                reportTextBox4.Text = "  DRIVER NO 3";
                reportTextBox5.Text = "  DRIVER NO 4";
                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = driverSN_Box1.Text;
                reportSerialNo3.Text = driverSN_Box2.Text;
                reportSerialNo4.Text = driverSN_Box3.Text;
                reportSerialNo5.Text = driverSN_Box4.Text;
            }
            else if (productButton10.Checked) //360X and XL manifolds
            {
                reportTextBox1.Text = "  MANIFOLD";
                reportTextBox2.Text = "  DRIVER NO 1";
                reportTextBox3.Text = "  DRIVER NO 2";
                reportTextBox4.Text = "  DRIVER NO 3";
                reportTextBox5.Text = "  DRIVER NO 4";
                reportSerialNo1.Text = SN_Box1.Text;
                reportSerialNo2.Text = driverSN_Box1.Text;
                reportSerialNo3.Text = driverSN_Box2.Text;
                reportSerialNo4.Text = driverSN_Box3.Text;
                reportSerialNo5.Text = driverSN_Box4.Text;
            }
            else if (productButton22.Checked) //360XL-MID
            {
                reportTextBox1.Text = "  SYSTEM";
                reportSerialNo1.Text = SN_Box1.Text;
                reportTextBox2.Text = "  DRIVER NO 1";

                reportSerialNo2.Text = driverSN_Box1.Text;

                if (driverSN_Box2.Visible)
                {
                    reportTextBox3.Text = "  DRIVER NO 2";
                    reportSerialNo3.Text = driverSN_Box2.Text;
                }

            }
            else
            {
                reportTextBox1.Text = "  SYSTEM";
                reportSerialNo1.Text = SN_Box1.Text;
            }
            /*******************************************************************************************************************************************************************************/
            /* 2nd sesion: Includede Cables/Assessories 
            /*******************************************************************************************************************************************************************************/
            if ((productButton1.Checked) || (productButton2.Checked) || (productButton3.Checked) || (productButton4.Checked) || (DUT.productModel == "100X") || (DUT.productModel == "100X-NAVY-V1") ||
              (productButton5.Checked) || (productButton7.Checked) || (productButton8.Checked) || (productButton9.Checked) || (productButton11.Checked))
            {
                reportAccessoryNo1.Text = "  Control Unit/MP3 Cable";

                if ((DUT.productModel == "300XRA") || (DUT.productModel == "450XL-RA") || (productButton5.Checked) ||
                  (productButton5.Checked) || (productButton8.Checked) || (productButton9.Checked) || (productButton11.Checked))
                {
                    reportAccessoryNo2.Text = "  Audio Cable";

                    if ((productButton5.Checked) || (productButton8.Checked) || (productButton9.Checked))
                    {
                        reportAccessoryNo3.Text = "  AC Power Supply Cable";
                    }
                }
                else if ((DUT.productModel == "100X") || (DUT.productModel == "100X-NAVY-V1"))
                {
                    reportAccessoryNo2.Text = "  Battery Pack";
                }
                else if (DUT.productModel == "1000")
                {
                    reportAccessoryNo2.Text = "  AUX Cable";
                    reportAccessoryNo3.Text = "  AC PS To DC Input Cable";
                }
                else if (DUT.productModel == "1000Xi")
                {
                    reportAccessoryNo2.Text = "  AC Power Supply Cable";
                    reportAccessoryNo3.Text = "  Short (Remote) MP3 cable";
                }
            }
            else if (DUT.productModel == "1000X2U")
            {
                reportAccessoryNo1.Text = "  Audio Cable";
            }
            else if (DUT.productModel == "100X-NAVY")
            {
                reportAccessoryNo1.Text = "  IPOD Touch";
                reportAccessoryNo2.Text = "  Battery Pack";
            }
        }

        private Boolean MissingSN()
        {
            Boolean notAllGood = false;

            if (troubleShootingCheckBox.Checked) return notAllGood;

            if (oper_Ini.Text == "")
            {
                DialogResult dialogResult = MessageBox.Show("'Operator Ini' is missing!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                notAllGood = true;
                return notAllGood;
            }

            if (WO_No.Text == "")
            {
                DialogResult dialogResult = MessageBox.Show("'WO Number' is missing!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                notAllGood = true;
                return notAllGood;
            }

            if ((SN_Box1.Visible) && (SN_Box1.Enabled) && (SN_Box1.Text == "")) notAllGood = true;
            if ((SN_Box2.Visible) && (SN_Box2.Enabled) && (SN_Box2.Text == "")) notAllGood = true;
            if ((SN_Box3.Visible) && (SN_Box3.Enabled) && (SN_Box3.Text == "")) notAllGood = true;
            if ((SN_Box4.Visible) && (SN_Box4.Enabled) && (SN_Box4.Text == "")) notAllGood = true;
            if ((SN_Box5.Visible) && (SN_Box5.Enabled) && (SN_Box5.Text == "")) notAllGood = true;
            if ((SN_Box6.Visible) && (SN_Box6.Enabled) && (SN_Box6.Text == "")) notAllGood = true;
            if ((SN_Box7.Visible) && (SN_Box7.Enabled) && (SN_Box7.Text == "")) notAllGood = true;
            if ((SN_Box8.Visible) && (SN_Box8.Enabled) && (SN_Box8.Text == "")) notAllGood = true;
            if ((SN_Box9.Visible) && (SN_Box9.Enabled) && (SN_Box9.Text == "")) notAllGood = true;
            if ((SN_Box10.Visible) && (SN_Box10.Enabled) && (SN_Box10.Text == "")) notAllGood = true;

            if ((driverSN_Box1.Visible) && (driverSN_Box1.Enabled) && (driverSN_Box1.Text == "")) notAllGood = true;
            if ((driverSN_Box2.Visible) && (driverSN_Box2.Enabled) && (driverSN_Box2.Text == "")) notAllGood = true;
            if ((driverSN_Box3.Visible) && (driverSN_Box3.Enabled) && (driverSN_Box3.Text == "")) notAllGood = true;
            if ((driverSN_Box4.Visible) && (driverSN_Box4.Enabled) && (driverSN_Box4.Text == "")) notAllGood = true;
            if ((driverSN_Box5.Visible) && (driverSN_Box5.Enabled) && (driverSN_Box5.Text == "")) notAllGood = true;
            if ((driverSN_Box6.Visible) && (driverSN_Box6.Enabled) && (driverSN_Box6.Text == "")) notAllGood = true;
            if ((driverSN_Box7.Visible) && (driverSN_Box7.Enabled) && (driverSN_Box7.Text == "")) notAllGood = true;
            if ((driverSN_Box8.Visible) && (driverSN_Box8.Enabled) && (driverSN_Box8.Text == "")) notAllGood = true;

            if (notAllGood)
            {
                DialogResult dialogResult = MessageBox.Show("Please enter the missing S/N(s)!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            return notAllGood;
        }

        private string assignAccessoryType()
        {
            string AccessoryType = null;

            if ((productButton5.Checked) || (productButton8.Checked) || (productButton9.Checked)) //1000X, 1000XVB, 1950XL
            {
                if (modelButton0.Checked)
                {
                    AccessoryType = "20FT-Cable";
                }
                else if (modelButton1.Checked)
                {
                    AccessoryType = "30FT-Cable";
                }
                else if (modelButton2.Checked)
                {
                    AccessoryType = "100FT-Cable";
                }
            }
            else if (productButton12.Checked) //DS60-X /w Transformers
            {
                if (testSelButton0.Checked)
                {
                    if (modelButton0.Checked)
                    {
                        AccessoryType = "70V-60W";
                    }
                    else if (modelButton1.Checked)
                    {
                        AccessoryType = "100V-60W";
                    }
                    else if (modelButton2.Checked)
                    {
                        AccessoryType = "25V-80W";
                    }
                    else if (modelButton3.Checked)
                    {
                        AccessoryType = "70V-80W";
                    }
                    else if (modelButton4.Checked)
                    {
                        AccessoryType = "100V-80W";
                    }
                }
                if (testSelButton1.Checked)
                {
                    if (modelButton0.Checked)
                    {
                        AccessoryType = "70V-40W";
                    }
                    else if (modelButton1.Checked)
                    {
                        AccessoryType = "100V-40W";
                    }
                    else if (modelButton2.Checked)
                    {
                        AccessoryType = "25V-60W";
                    }
                    else if (modelButton3.Checked)
                    {
                        AccessoryType = "70V-60W";
                    }
                    else if (modelButton4.Checked)
                    {
                        AccessoryType = "100V-60W";
                    }
                }
                if (testSelButton2.Checked)
                {
                    if (modelButton0.Checked)
                    {
                        AccessoryType = "70V-20W";
                    }
                    else if (modelButton1.Checked)
                    {
                        AccessoryType = "100V-20W";
                    }
                    else if (modelButton2.Checked)
                    {
                        AccessoryType = "25V-40W";
                    }
                    else if (modelButton3.Checked)
                    {
                        AccessoryType = "70V-40W";
                    }
                    else if (modelButton4.Checked)
                    {
                        AccessoryType = "100V-40W";
                    }
                }
            }
            else if (productButton13.Checked) //DS60-XL /w Transformers
            {
                if (testSelButton0.Checked)
                {
                    if (modelButton0.Checked)
                    {
                        AccessoryType = "70V-160W";
                    }
                    else if (modelButton1.Checked)
                    {
                        AccessoryType = "100V-160W";
                    }
                }
                else if (testSelButton1.Checked)
                {
                    if (modelButton0.Checked)
                    {
                        AccessoryType = "70V-120W";
                    }
                    else if (modelButton1.Checked)
                    {
                        AccessoryType = "100V-120W";
                    }
                }
                else if (testSelButton2.Checked)
                {
                    if (modelButton0.Checked)
                    {
                        AccessoryType = "70V-80W";
                    }
                    else if (modelButton1.Checked)
                    {
                        AccessoryType = "100V-80W";
                    }
                }
            }
            return AccessoryType;

        }

        private string assignFileNameForMaxOutput()
        {
            string sampleNameVOut = null;

            string AcessoryType = assignAccessoryType();

            if (testSelButton0.Checked)
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameVOut = "\\Screening Test 0 (SPL) - ";
                }
                else if ((DUT.productModel == "1000X2U") || (DUT.productModel == "1000") || (DUT.productModel == "100X") || (productButton22.Checked))
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL - ";
                }
                else if (DUT.productModel == "100X-NAVY")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (MP3-2, MaxPwr-On, VB-Off) - ";
                }
                //DS60X & DS60XL w transformer
                else if ((productButton12.Checked) || (productButton13.Checked))
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + ", " + AcessoryType + "] Maximum SPL (" + AcessoryType + ") - ";
                }
                else if ((DUT.productModel == "SS100") || (DUT.productModel == "SS300"))
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (70V-60W Terminal) - ";
                }
                else if (DUT.productModel == "SSX")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (100V-120W Terminal) - ";
                }
                else if (DUT.productModel == "SSX wo Trans")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL - ";
                }
                else if (DUT.productModel == "SSX60")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (100V-60W Terminal) - ";
                }
                else if (DUT.productModel == "SSX60 wo Trans")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (60W) - ";
                }
                else if (DUT.productModel == "360X Manifold")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (320W) - ";
                }
                else if (DUT.productModel == "360XL Manifold")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (640W) - ";
                }
                else if (DUT.productModel == "SS400")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (360W at 4.0 Ohms) - ";
                }
                //360Xm with Trans --> Two input powers
                else if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W"))
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (HIGH Power-TOP) - ";
                }
                //360Xm with Trans --> One input power
                else if (productButton20.Checked)
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (HIGH Power) - ";
                } //360Xm 1-ST wo Trans
                  // 360Xm wo Trans-- > One input power
                else if (DUT.productModel == "360Xm 1-ST wo Trans")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL - ";
                }
                // 360Xm wo Trans-- > Two input power
                else if (DUT.productModel == "360Xm 2-ST wo Trans")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (TOP) - ";
                } //
                else
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (Wide-Off) - ";
                }
            }
            else if (testSelButton1.Checked)
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameVOut = "\\Screening Test 1 (SPL) - ";
                }
                else if (productButton18.Checked == true)
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL - ";
                } //LRAD-RX
                else if (DUT.productModel == "SS400")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Accelerometer Output (With 1kHz Tone) - ";
                }
                // 360Xm wo Trans-- > Two input power
                else if (DUT.productModel == "360Xm 2-ST wo Trans")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (BOT) - ";
                } //
            }
            else if (testSelButton2.Checked) //Alert tone for RXs
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameVOut = "\\Screening Test 2 (SPL) - ";
                }
                else if (productButton18.Checked == true)
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Alert Tone SPL - ";
                } //LRAD-RX               
            }
            else if (testSelButton3.Checked) //Max SPL for 2-ST 360Xm
            {
                if ((troubleShootingCheckBox.Checked) && ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W")))
                {
                    sampleNameVOut = "\\Screening Test 3 (SPL) - ";
                }
                else if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W"))
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (HIGH Power-BOT) - ";
                }
            }
            else if (testSelButton7.Checked)
            {
                if ((troubleShootingCheckBox.Checked) && (DUT.productModel == "100X-NAVY"))
                {
                    sampleNameVOut = "\\Screening Test 7 (SPL) - ";
                }
                else if (DUT.productModel == "100X-NAVY")
                {
                    sampleNameVOut = "\\[" + DUT.SN[1] + "] Maximum SPL (IPOD, MaxPwr-On, VB-Off) - ";
                }
            }

            return sampleNameVOut;
        }

        private string assignFileNameForFR()
        {
            string sampleNameFR = null;

            string AcessoryType = assignAccessoryType();

            if (testSelButton0.Checked)
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameFR = "\\Screening Test 0 - ";
                }
                else if ((DUT.productModel == "100X") || (DUT.productModel == "100X-NAVY-V1"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Max Vol) - ";
                }
                else if (DUT.productModel == "100X-NAVY")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (MP3-1, MaxPwr-On, VB-Off) - ";
                }
                else if (DUT.productModel == "1000X")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Max Output Power, " + AcessoryType + ") - ";
                }
                else if (DUT.productModel == "1000XVB")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Max Output Power, VB=Off, " + AcessoryType + ") - ";
                }
                else if (DUT.productModel == "1950XL")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Wide-Off, " + AcessoryType + ") - ";
                }
                else if (DUT.productModel == "1000X2U")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Max Vol) - ";
                }
                else if (DUT.productModel == "1000")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Wide-Off) - ";
                }
                //DS60X & DS60XL w transformer
                else if ((productButton12.Checked) || (productButton13.Checked))
                {
                    sampleNameFR = "\\[Hd=" + DUT.SN[1] + ", Transformer Box=" + DUT.SN[2] + "] Sensitivity (" + AcessoryType + ") - ";
                }
                else if ((DUT.productModel == "SS400") || (DUT.productModel == "DS60-X") || (DUT.productModel == "SSX wo Trans") || (DUT.productModel == "SSX60 wo Trans"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] " + "Frequency Response (Sensitivity) - ";
                }
                else if (DUT.productModel == "SSX wo Trans")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity - ";
                }
                else if (DUT.productModel == "SSX60")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (100V-60W Terminal) - ";
                }
                else if ((DUT.productModel == "SS100") || (DUT.productModel == "SS300"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (70V-60W Terminal) - ";
                }
                else if (DUT.productModel == "SSX")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (100V-120W Terminal) - ";
                }
                else if (DUT.productModel == "SSX60 wo Trans")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity - ";
                }
                else if (DUT.productModel == "360X Manifold")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Freq Res - ";
                }
                else if (DUT.productModel == "360XL Manifold")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Freq Res - ";
                }
                else if (productButton18.Checked)
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Frequency Sweep (300Hz - 15 kHz) - ";
                } //LRAD-RX 
                else if (productButton22.Checked)
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Frequency Response - ";
                }
                //360Xm 2-ST with Trans --> Two input powers
                else if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (HIGH Power-TOP) - ";
                }
                //360Xm with Trans --> One input power
                else if ((DUT.productModel == "360Xm 1-ST 25V-60W") || (DUT.productModel == "360Xm 1-ST 100V-60W") ||
                  (DUT.productModel == "360Xm 2-ST 100V-120W") || (DUT.productModel == "360Xm 4-ST 100V-240W"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (HIGH Power) - ";
                }
                //360Xm 1-ST wo Trans --> One input power
                else if (DUT.productModel == "360Xm 1-ST wo Trans")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity - ";
                }
                //360Xm 2-ST wo Trans --> Two input power
                else if (DUT.productModel == "360Xm 2-ST wo Trans")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (TOP) - ";
                }
                else
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Wide-Off) - ";
                }

            }
            else if (testSelButton1.Checked)
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameFR = "\\Screening Test 1 - ";
                }
                //DS60X & DS60XL w transformer                
                else if ((productButton12.Checked) || (productButton13.Checked))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (" + AcessoryType + ") - ";
                }
                else if (DUT.productModel == "SSX")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (100V-60W Terminal) - ";
                }
                else if (DUT.productModel == "SSX60")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (100V-40W Terminal) - ";
                }
                else if ((DUT.productModel == "SS100") || (DUT.productModel == "SS300"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (70V-40W Terminal) - ";
                }
                //360Xm 2-ST with Trans --> Two input powers
                else if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (MID Power-TOP) - ";
                }
                //360Xm with Trans --> One input power
                else if ((DUT.productModel == "360Xm 1-ST 25V-60W") || (DUT.productModel == "360Xm 1-ST 100V-60W") ||
                  (DUT.productModel == "360Xm 2-ST 100V-120W") || (DUT.productModel == "360Xm 4-ST 100V-240W"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (MID Power) - ";
                }
                //360Xm 2-ST wo Trans --> Two input power
                else if (DUT.productModel == "360Xm 2-ST wo Trans")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (BOT) - ";
                }
                else
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Vol Ctrl Knob Function) - ";
                }
            }
            else if (testSelButton2.Checked)
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameFR = "\\Screening Test 2 - ";
                }
                else if (DUT.productModel == "1000X2U")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Mute Function - ";
                }
                else if ((productButton12.Checked) || (productButton13.Checked))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (" + AcessoryType + ") - ";
                }
                else if (DUT.productModel == "SSX")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (100V-30W Terminal) - ";
                }
                else if (DUT.productModel == "SSX60")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (100V-20W Terminal) - ";
                }
                else if ((DUT.productModel == "SS100") || (DUT.productModel == "SS300"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (70V-20W Terminal) - ";
                }
                else if (DUT.productModel == "1000")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (MIC input, Wide-Off, Max Vol) - ";
                }
                //360Xm 2-ST with Trans --> Two input powers
                else if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (LOW Power-TOP) - ";
                }
                //360Xm with Trans --> One input power
                else if ((DUT.productModel == "360Xm 1-ST 25V-60W") || (DUT.productModel == "360Xm 1-ST 100V-60W") ||
                  (DUT.productModel == "360Xm 2-ST 100V-120W") || (DUT.productModel == "360Xm 4-ST 100V-240W"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (LOW Power) - ";
                }
                else
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (MIC Input) - ";
                }
            }
            else if (testSelButton3.Checked)
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameFR = "\\Screening Test 3 - ";
                }
                else if (DUT.productModel == "1000X")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Low Output Power, " + AcessoryType + ") - ";
                }
                else if (DUT.productModel == "1000XVB")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Max Output Power, VB=On, " + AcessoryType + ") - ";
                }
                else if (DUT.productModel == "100X-NAVY")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (MP3-2, MaxPwr-On, VB-On) - ";
                }
                else if (DUT.productModel == "1000")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Wide-On) - ";
                }
                //360Xm 2-ST with Trans --> Two input powers
                else if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (HIGH Power-BOT) - ";
                }
                else
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Wide-On) - ";
                }
            }
            else if (testSelButton4.Checked)
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameFR = "\\Screening Test 4 - ";
                }
                else if (DUT.productModel == "1000XVB")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Low Output Power, VB=Off, " + AcessoryType + ") - ";
                }
                else if (DUT.productModel == "100X-NAVY")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (MP3-2, MaxPwr-On, VB-Off) - ";
                }
                else if (DUT.productModel == "1000")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Narrow-Off) - ";
                }
                //360Xm 2-ST with Trans --> Two input powers
                else if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (MID Power-BOT) - ";
                }
                else
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Narrow-Off) - ";
                }
            }
            else if (testSelButton5.Checked)
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameFR = "\\Screening Test 5 - ";
                }
                else if (DUT.productModel == "100X-NAVY")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (MP3-2, MaxPwr-On, VB-Off) - ";
                }
                else if (DUT.productModel == "1000")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Narrow-On) - ";
                }
                //360Xm 2-ST with Trans --> Two input powers
                else if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W"))
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (LOW Power-BOT) - ";
                }
                else
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (Narrow-On) - ";
                }
            }
            else if (testSelButton6.Checked)
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameFR = "\\Screening Test 6 - ";
                }
                else if (DUT.productModel == "1000X2U")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Output Noise Level - ";
                }
                else if (DUT.productModel == "1000")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Output Noise Level - ";
                }
                else
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Output Noise Level - ";
                }
            }
            else if (testSelButton7.Checked)
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameFR = "\\Screening Test 7 - ";
                }
                else if (DUT.productModel == "100X-NAVY")
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Sensitivity (IPOD, MaxPwr-On, VB-Off) - ";
                }
                else
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] Limited Key Switch Function - ";
                }
            }
            else if (testSelButton9.Checked)
            {
                if (troubleShootingCheckBox.Checked)
                {
                    sampleNameFR = "\\Screening Test 9 (Tip) - ";
                }
                else
                {
                    sampleNameFR = "\\[" + DUT.SN[1] + "] AUX Cable (Tip) - ";
                }
            }
            return sampleNameFR;
        }

        private void De_activateUserInputs(bool unLock)
        {
            if (unLock)
            {
                //   SN1_Box.Enabled = false; SN2_Box.Enabled = false;
                //   SN3_Box.Enabled = false; 
                groupBox1.Enabled = false;
                groupBox2.Enabled = false;
                groupBox3.Enabled = false;
                //  groupBox4.Enabled = false; groupBox5.Enabled = false; groupBox6.Enabled = false; groupBox10.Enabled = false;
                groupBox4.Enabled = false;
                groupBox5.Enabled = false;
                groupBox10.Enabled = false;
                remarkBox.Enabled = false;
            }
            else
            {
                //  serialBox1.Enabled = true; serialBox2.Enabled = true;
                //   startButton.Enabled = true; printButton.Enabled = true;
                groupBox1.Enabled = true;
                groupBox2.Enabled = true;
                groupBox3.Enabled = true;
                //   groupBox4.Enabled = true; groupBox5.Enabled = true; groupBox6.Enabled = true; groupBox10.Enabled = true;
                groupBox4.Enabled = true;
                groupBox5.Enabled = true;
                groupBox10.Enabled = true;
                remarkBox.Enabled = true;

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Set_LRAD_MP3_LowerMic();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Set_LRADX_MP3_UpperMic();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Set_LRADX_MP3_LowerMic();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Set_LRADX_MIC_UpperMic();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Set_LRADX_MIC_LowerMic();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Set_LRAD_MIC_LowerMic();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Set_SS_UpperMic();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Set_SS_LowerMic();
        }

        private void MP3_2_100XNavy_Click(object sender, EventArgs e)
        {
            Set_100XNavy_MP3_2_LRAD_Mic();
        }

        private void IPOD_100XNavy_Click(object sender, EventArgs e)
        {
            Set_100XNavy_IPOD_LRAD_Mic();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Set_360XManifolds_LowerMic();
        }

        private void Set_360XManifolds_LowerMic()
        {
            DUT.portCState = 0x00;
            channelSetup(0x2B, 0x2B, DUT.portCState);
        }
        private void Set_100XNavy_MP3_2_LRAD_Mic()
        {
            DUT.portCState = 0x00;
            channelSetup(0xaa, 0xaa, DUT.portCState);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Set_LRADX_LowerMic_AllLinesRemoved();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Set_LRADX_UpperMic_AllLinesRemoved();
        }

        private void Set_LRADX_UpperMic_AllLinesRemoved()
        {
            DUT.portCState = 0x00;
            channelSetup(0x4a, 0x4a, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "/LRADX Upper MIC selected. All lines removed";
                debugLabel4.Update();
            }
        }

        private void Set_LRADX_LowerMic_AllLinesRemoved()
        {
            DUT.portCState = 0x00;
            channelSetup(0x2a, 0x2a, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "/LRADX Lower MIC selected. All lines removed";
                debugLabel4.Update();
            }
        }

        private void Set_100XNavy_IPOD_LRAD_Mic()
        {
            DUT.portCState = 0x00;
            channelSetup(0x2a, 0x3a, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "/100X NAVY IPOD input";
                debugLabel4.Update();
            }
        }

        private void Set_LRADX_MP3_UpperMic()
        {
            DUT.portCState = 0x00;
            channelSetup(0x68, 0x68, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "/LRADX Upper MIC selected. MP3 Input";
                debugLabel4.Update();
            }
        }

        private void Set_LRADX_MP3_LowerMic()
        {
            DUT.portCState = 0x00;
            channelSetup(0x28, 0x28, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "/LRADX Lower MIC selected. MP3 Input";
                debugLabel4.Update();
            }
        }

        private void Set_LRADX_MIC_UpperMic()
        {
            DUT.portCState = 0x00;
            channelSetup(0x6e, 0x6e, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "/LRADX Upper MIC selected. MIC Input";
                debugLabel4.Update();
            }
        }

        private void Set_LRADX_MIC_LowerMic()
        {
            DUT.portCState = 0x00;
            channelSetup(0x6e, 0x2e, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "/LRADX Lower MIC selected. MIC Input";
                debugLabel4.Update();
            }
        }

        private void Set_LRAD_MP3_LowerMic()
        {
            DUT.portCState = 0x00;
            channelSetup(0x22, 0x22, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "/LRAD Lower MIC selected. MP3 Input";
                debugLabel4.Update();
            }
        }

        private void Set_LRAD_MIC_LowerMic()
        {
            DUT.portCState = 0x00;
            channelSetup(0x3a, 0x3a, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "/LRAD Lower MIC selected. MIC Input";
                debugLabel4.Update();
            }
        }

        private void Set_SS_UpperMic()
        {
            DUT.portCState = 0x00;
            channelSetup(0x4b, 0x4b, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "/SS Upper MIC selected.              ";
                debugLabel4.Update();
            }
        }

        private void Set_SS_LowerMic()
        {
            DUT.portCState = 0x00;
            channelSetup(0x2b, 0x2b, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "/SS Lower MIC selected.              ";
                debugLabel4.Update();
            }
        }

        private void TransHI_PWR_Click(object sender, EventArgs e)
        {
            if ((DUT.productModel == "SS300") || (productButton20.Checked)) //Using test cart /w trans selectable outputs
            {
                Set_Trans_HiPWR();
            }
            else
            {
                Set_LRADX_MP3_LowerMic();
            }
        }

        private void Set_Trans_HiPWR()
        {
            DUT.portCState = (byte)(DUT.portCState | 8);
            channelSetup(0x01, 0xf7, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "Using test cart /w trans selectable outputs --> HI Power line selected";
                debugLabel4.Update();
            }
        }

        private void TransMID_PWR_Click(object sender, EventArgs e)
        {
            if ((DUT.productModel == "SS300") || (productButton20.Checked)) //Using test cart /w trans selectable outputs
            {
                Set_Trans_MIDPWR();
            }
            else
            {
                Set_LRADX_MP3_LowerMic();
            }
        }

        private void Set_Trans_MIDPWR()
        {
            DUT.portCState = (byte)(DUT.portCState | 8);
            channelSetup(0x02, 0xf7, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "Using test cart /w trans selectable outputs --> MID Power line selected";
                debugLabel4.Update();
            }
        }

        private void TransLOW_PWR_Click(object sender, EventArgs e)
        {
            if ((DUT.productModel == "SS300") || (productButton20.Checked)) //Using test cart /w trans selectable outputs
            {
                Set_Trans_LOWPWR();
            }
            else
            {
                Set_LRADX_MP3_LowerMic();
            }
        }

        private void Set_Trans_LOWPWR()
        {
            DUT.portCState = (byte)(DUT.portCState | 8);
            channelSetup(0x04, 0xf7, DUT.portCState);

            if (DUT.Debug == 1)
            {
                debugLabel4.Text = "Using test cart /w trans selectable outputs --> LOW Power line selected";
                debugLabel4.Update();
            }
        }

        /******************************************************************************************************/
        /* Activate a test sequence from AP sequencer and setup Ref limits                                    */
        /******************************************************************************************************/
        private Boolean SetupSequenceSelect(string APOutput, string sequenceID, string refFileName, string refFileNamePlus, string refFileNameMinus)
        {
            string refCurvePlus = Folder.refFileFolder + refFileNamePlus;
            string refCurveMinus = Folder.refFileFolder + refFileNameMinus;
            string refCurve = Folder.refFileFolder + refFileName;

            Boolean BadStatus = false;
            Boolean NoUpperRefCurve = false;

            if (!File.Exists(refCurvePlus))
            {
                NoUpperRefCurve = true;
            }

            //set the checked state in Sequence mode
            APx.Sequence.GetMeasurement("Signal Path1", "Signal Path Setup").Checked = false;
            APx.Sequence.GetMeasurement("Signal Path1", "Acoustic Response").Checked = false;
            APx.Sequence.GetMeasurement("Signal Path1", "Signal Analyzer").Checked = false;
            APx.Sequence.GetMeasurement("Signal Path1", "Stepped Frequency Sweep").Checked = false;

            /////////////////////////////////////////////////////////////////////////////////////////
            /* General APx configuration */
            APx.SignalPathSetup.AnalogInput.ChannelCount = 1;
            APx.SignalPathSetup.AnalogOutput.ChannelCount = 2;

            ///////////////////////////////////////////////////////////////////////////////////////////        
            if ((DUT.productModel == "SS400") && (testSelButton1.Checked))
            {
                APx.SignalPathSetup.AnalogInput.SingleInputChannel = SingleInputChannelIndex.Ch2;
                APx.SignalPathSetup.AcousticInput = false;
            }
            else
            {
                APx.SignalPathSetup.AnalogInput.SingleInputChannel = SingleInputChannelIndex.Ch1;
                APx.SignalPathSetup.AcousticInput = true;
            }

            if (sequenceID == "FR")
            {
                try
                {
                    APx.ShowMeasurement("Signal Path1", "Acoustic response");
                    APx.Sequence.GetMeasurement("Signal Path1", "Acoustic response").Checked = true;

                    APx.ActiveMeasurement.Graphs[1].Checked = false;
                    APx.ActiveMeasurement.Graphs[2].Checked = false;

                    APx.ActiveMeasurement.Graphs[3].Checked = true;
                    APx.ActiveMeasurement.Graphs[3].Name = DUT.graphTitle;
                    APx.ActiveMeasurement.Graphs[3].Show();

                    if (APOutput == "MP3_Ch1")
                    {
                        if (APx.AcousticResponse.GeneratorWithPilot.GetChannelEnabled(OutputChannelIndex.Ch2))
                        {
                            APx.AcousticResponse.GeneratorWithPilot.SetChannelEnabled(OutputChannelIndex.Ch2, false);
                            APx.AcousticResponse.GeneratorWithPilot.SetChannelEnabled(OutputChannelIndex.Ch1, true);

                            if (DUT.Debug == 1)
                            {
                                debugLabel4.Text = "LRADX MP3 input --> AP Ch1 Output selected";
                                debugLabel4.Update();
                            }
                        }
                    }
                    else if (APOutput == "MIC_Ch2")
                    {
                        if (APx.AcousticResponse.GeneratorWithPilot.GetChannelEnabled(OutputChannelIndex.Ch1))
                        {
                            APx.AcousticResponse.GeneratorWithPilot.SetChannelEnabled(OutputChannelIndex.Ch1, false);
                            APx.AcousticResponse.GeneratorWithPilot.SetChannelEnabled(OutputChannelIndex.Ch2, true);

                            if (DUT.Debug == 1)
                            {
                                debugLabel4.Text = "LRADX MIC input --> AP Ch2 Output selected";
                                debugLabel4.Update();
                            }
                        }
                    }

                    if ((testSelButton2.Checked) && !(productButton6.Checked) && !(productButton12.Checked) && !(productButton13.Checked) && !(DUT.productModel == "SS100") &&
                      !(DUT.productModel == "SS300") && !(productButton16.Checked) && !(productButton20.Checked) && (DUT.noBlackBoxSetup == 1))
                    {
                        //    APx.SignalPathSetup.AnalogOutput.ChannelCount = 2;
                        APx.AcousticResponse.GeneratorWithPilot.SetChannelEnabled(OutputChannelIndex.Ch1, false);
                        APx.AcousticResponse.GeneratorWithPilot.SetChannelEnabled(OutputChannelIndex.Ch2, true);

                        if (DUT.Debug == 1)
                        {
                            debugLabel4.Text = "LRADX MP3 input --> AP Ch2 Output selected";
                            debugLabel4.Update();
                        }
                    }
                    else
                    {
                        APx.AcousticResponse.GeneratorWithPilot.SetChannelEnabled(OutputChannelIndex.Ch1, true);

                        if (DUT.Debug == 1)
                        {
                            debugLabel4.Text = "LRADX MP3 input --> AP Ch1 Output selected";
                            debugLabel4.Update();
                        }
                    }

                    APx.AcousticResponse.GeneratorWithPilot.Levels.Sweep.SetValue(OutputChannelIndex.Ch1, DUT.stimulusFR);
                    APx.AcousticResponse.GeneratorWithPilot.Durations.Sweep.Value = 0.3;
                    APx.AcousticResponse.GeneratorWithPilot.Durations.PreSweep.Value = 0.05;

                    APx.AcousticResponse.ImpulseResponse.TimeWindowAutoStart = false;
                    APx.AcousticResponse.ImpulseResponse.TimeWindowStart.Value = DUT.WindowBegin;
                    APx.AcousticResponse.ImpulseResponse.TimeWindowEnd.Value = DUT.WindowEnd;

                    APx.AcousticResponse.GeneratorWithPilot.Frequencies.Start.Value = DUT.freqStart;
                    APx.AcousticResponse.GeneratorWithPilot.Frequencies.Stop.Value = DUT.freqStop;
                    APx.AcousticResponse.GeneratorWithPilot.EQSettings.EQTableType = EQType.None;

                    if (DUT.curveSmooth == "0")
                    {
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().OctaveSmoothing = OctaveSmoothingType.None;
                    }
                    else if (DUT.curveSmooth == "1")
                    {
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().OctaveSmoothing = OctaveSmoothingType.Octave1;
                    }
                    else if (DUT.curveSmooth == "3")
                    {
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().OctaveSmoothing = OctaveSmoothingType.Octave3;
                    }
                    else if (DUT.curveSmooth == "6")
                    {
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().OctaveSmoothing = OctaveSmoothingType.Octave6;
                    }
                    else if (DUT.curveSmooth == "9")
                    {
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().OctaveSmoothing = OctaveSmoothingType.Octave9;
                    }
                    else if (DUT.curveSmooth == "12")
                    {
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().OctaveSmoothing = OctaveSmoothingType.Octave12;
                    }
                    else if (DUT.curveSmooth == "24")
                    {
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().OctaveSmoothing = OctaveSmoothingType.Octave24;
                    }

                    IGraphCollection Graphs =
                      default(IGraphCollection);
                    Graphs = APx.ActiveMeasurement.Graphs;
                    IXYGraph XYGraph =
                      default(IXYGraph);

                    XYGraph = Graphs[3].Result.AsXYGraph();
                    XYGraph.YAxis.IsLog = false;
                    XYGraph.YAxis.Unit = "dBSPL";
                    XYGraph.XAxis.IsLog = true;

                    if (DUT.YAxisMax == -100)
                    {
                        XYGraph.YAxis.RangeType = GraphRangeType.Autoscale;
                    }
                    else
                    {
                        XYGraph.YAxis.RangeType = GraphRangeType.Fixed;
                        XYGraph.YAxis.Maximum = DUT.YAxisMax;
                        XYGraph.YAxis.Minimum = DUT.YAxisMin;
                    }

                    XYGraph.XAxis.RangeType = GraphRangeType.Fixed;
                    XYGraph.XAxis.Maximum = DUT.XAxisMax;
                    XYGraph.XAxis.Minimum = DUT.XAxisMin;

                    APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().UpperLimit.Clear();
                    APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().LowerLimit.Clear();

                    APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().UpperLimit.Enabled = true;
                    if (NoUpperRefCurve)
                    {
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().UpperLimit.ImportData(refCurve);
                    }
                    else
                    {
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().UpperLimit.ImportData(refCurvePlus);
                    }

                    APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().UpperLimit.OffsetValues(DUT.SweepUpperLimit);

                    if (DUT.curveSmooth == "MUTE")
                    {
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().LowerLimit.Enabled = false;
                    }
                    else
                    {
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().LowerLimit.Enabled = true;
                        if (NoUpperRefCurve)
                        {
                            APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().LowerLimit.ImportData(refCurve);
                        }
                        else
                        {
                            APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().LowerLimit.ImportData(refCurveMinus);
                        }
                        APx.ActiveMeasurement.Graphs[3].Result.AsSmoothResult().LowerLimit.OffsetValues(-DUT.SweepLowerLimit);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + " ---> " + refCurve + "\n PLEASE INFORM THE ENGINEER IN CHARGE!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    BadStatus = true;
                    DUT.hasResults = false;
                }
            }
            else if (sequenceID == "SPL")
            {
                try
                {
                    APx.ShowMeasurement("Signal Path1", "Signal Analyzer");
                    APx.Sequence.GetMeasurement("Signal Path1", "Signal Analyzer").Checked = true;

                    APx.SignalAnalyzer.Generator.Levels.SetValue(OutputChannelIndex.Ch1, DUT.stimulusSPL);
                    APx.ActiveMeasurement.Graphs[0].Show();
                    APx.ActiveMeasurement.Graphs[0].Checked = false; //FFT
                    APx.ActiveMeasurement.Graphs[2].Checked = false; //FFT --> Smooth for Noise meas.                                        

                    APx.ActiveMeasurement.Graphs[1].Show();
                    APx.ActiveMeasurement.Graphs[1].Checked = true; //FFT --> Max SPL
                    APx.ActiveMeasurement.Graphs[1].Name = DUT.SPLGraphTitle;

                    APx.ActiveMeasurement.Graphs[1].Result.AsStatisticsMeterFromXYResult().Function = MeterStatisticsFunctionType.Max;
                    APx.ActiveMeasurement.Graphs[1].Result.AsStatisticsMeterFromXYResult().Limits.Upper.Enabled = true;
                    APx.ActiveMeasurement.Graphs[1].Result.AsStatisticsMeterFromXYResult().Limits.Lower.Enabled = true;

                    var maxResult = APx.ActiveMeasurement.Graphs[1].Result.AsStatisticsMeterFromXYResult();
                    maxResult.Function = MeterStatisticsFunctionType.Max;
                    maxResult.Axis.RangeType = GraphRangeType.Fixed;
                    maxResult.Axis.Minimum = DUT.RefSPL - (DUT.RefSPL * 0.2);
                    maxResult.Axis.Maximum = DUT.RefSPL + (DUT.RefSPL * 0.2);

                    if ((DUT.productModel == "SS400") && (testSelButton1.Checked))
                    {
                        APx.ActiveMeasurement.Graphs[1].Result.AsStatisticsMeterFromXYResult().Limits.Upper.SetValue(InputChannelIndex.Ch1, DUT.RefSPL + DUT.SPLUpperLimit, "V");
                        APx.ActiveMeasurement.Graphs[1].Result.AsStatisticsMeterFromXYResult().Limits.Lower.SetValue(InputChannelIndex.Ch1, DUT.RefSPL - DUT.SPLLowerLimit, "V");
                    }
                    else
                    {
                        APx.ActiveMeasurement.Graphs[1].Result.AsStatisticsMeterFromXYResult().Limits.Upper.SetValue(InputChannelIndex.Ch1, DUT.RefSPL + DUT.SPLUpperLimit, "dBSPL");
                        APx.ActiveMeasurement.Graphs[1].Result.AsStatisticsMeterFromXYResult().Limits.Lower.SetValue(InputChannelIndex.Ch1, DUT.RefSPL - DUT.SPLLowerLimit, "dBSPL");
                    }

                    APx.SignalAnalyzer.Generator.Waveform = "Sine";
                    APx.SignalAnalyzer.Averages = 2;
                    APx.SignalAnalyzer.Generator.Frequency.Value = DUT.freqSPL;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + " ---> " + "\n PLEASE INFORM THE ENGINEER IN CHARGE!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    BadStatus = true;
                    DUT.hasResults = false;
                }

            }
            else if (sequenceID == "Noise")
            {
                try
                {
                    APx.ShowMeasurement("Signal Path1", "Signal Analyzer");
                    APx.Sequence.GetMeasurement("Signal Path1", "Signal Analyzer").Checked = true;

                    APx.SignalAnalyzer.Generator.Levels.SetValue(OutputChannelIndex.Ch1, DUT.stimulusSPL);
                    APx.ActiveMeasurement.Graphs[0].Checked = false; //FFT
                    APx.ActiveMeasurement.Graphs[1].Checked = false; //FFT --> Max SPL    

                    APx.ActiveMeasurement.Graphs[2].Checked = true; //FFT --> Smooth for Noise meas.                                                
                    APx.ActiveMeasurement.Graphs[2].Name = DUT.graphTitle;
                    APx.ActiveMeasurement.Graphs[2].Show();
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    APx.SignalAnalyzer.Averages = 5;

                    IGraphCollection Graphs =
                      default(IGraphCollection);
                    Graphs = APx.ActiveMeasurement.Graphs;
                    IXYGraph XYGraph =
                      default(IXYGraph);

                    XYGraph = Graphs[2].Result.AsXYGraph();
                    XYGraph.YAxis.IsLog = false;
                    XYGraph.YAxis.Unit = "dBSPL";
                    XYGraph.XAxis.IsLog = true;

                    if (DUT.YAxisMax == -100)
                    {
                        XYGraph.YAxis.RangeType = GraphRangeType.Autoscale;
                    }
                    else
                    {
                        XYGraph.YAxis.RangeType = GraphRangeType.Fixed;
                        XYGraph.YAxis.Maximum = DUT.YAxisMax;
                        XYGraph.YAxis.Minimum = DUT.YAxisMin;
                    }

                    XYGraph.XAxis.RangeType = GraphRangeType.Fixed;
                    XYGraph.XAxis.Maximum = DUT.XAxisMax;
                    XYGraph.XAxis.Minimum = DUT.XAxisMin;
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////                    
                    /*      APx.ActiveMeasurement.Graphs[1].Result.AsStatisticsMeterFromMeterResult().Checked = true;
                          APx.ActiveMeasurement.Graphs[1].Result.AsStatisticsMeterFromMeterResult().Limits.Upper.Enabled = true;
                          APx.ActiveMeasurement.Graphs[1].Result.AsStatisticsMeterFromMeterResult().Limits.Upper.SetValue(InputChannelIndex.Ch1, DUT.RefSPL + DUT.SPLUpperLimit, "dBSPL");
                          APx.SignalAnalyzer.Generator.Waveform = "Sine";
                          APx.SignalAnalyzer.Averages = 5;
                          APx.SignalAnalyzer.Generator.Frequency.Value = DUT.freqSPL;  */

                    APx.ActiveMeasurement.Graphs[2].Result.AsSmoothResult().OctaveSmoothing = OctaveSmoothingType.None;
                    APx.ActiveMeasurement.Graphs[2].Result.AsSmoothResult().UpperLimit.Clear();
                    APx.ActiveMeasurement.Graphs[2].Result.AsSmoothResult().UpperLimit.Enabled = true;
                    APx.ActiveMeasurement.Graphs[2].Result.AsSmoothResult().LowerLimit.Enabled = false;
                    APx.ActiveMeasurement.Graphs[2].Result.AsSmoothResult().UpperLimit.ImportData(refCurve);
                    APx.ActiveMeasurement.Graphs[2].Result.AsSmoothResult().UpperLimit.OffsetValues(DUT.SweepUpperLimit);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + " ---> " + refCurve + "\n PLEASE INFORM THE ENGINEER IN CHARGE!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    BadStatus = true;
                    DUT.hasResults = false;
                }

            }
            else if (sequenceID == "FreqSweep")
            {
                try
                {
                    APx.ShowMeasurement("Signal Path1", "Stepped Frequency Sweep");
                    APx.Sequence.GetMeasurement("Signal Path1", "Stepped Frequency Sweep").Checked = true;

                    APx.ActiveMeasurement.Graphs[0].Checked = true;
                    APx.ActiveMeasurement.Graphs[0].Name = DUT.graphTitle;
                    APx.ActiveMeasurement.Graphs[0].Show();

                    IGraphCollection Graphs =
                      default(IGraphCollection);
                    Graphs = APx.ActiveMeasurement.Graphs;
                    IXYGraph XYGraph =
                      default(IXYGraph);

                    XYGraph = Graphs[0].Result.AsXYGraph();
                    XYGraph.YAxis.IsLog = false;
                    XYGraph.YAxis.Unit = "dBSPL";
                    XYGraph.XAxis.IsLog = true;

                    if (DUT.YAxisMax == -100)
                    {
                        XYGraph.YAxis.RangeType = GraphRangeType.Autoscale;
                    }
                    else
                    {
                        XYGraph.YAxis.RangeType = GraphRangeType.Fixed;
                        XYGraph.YAxis.Maximum = DUT.YAxisMax;
                        XYGraph.YAxis.Minimum = DUT.YAxisMin;
                    }

                    XYGraph.XAxis.RangeType = GraphRangeType.Fixed;
                    XYGraph.XAxis.Maximum = DUT.XAxisMax;
                    XYGraph.XAxis.Minimum = DUT.XAxisMin;

                    APx.SteppedFrequencySweep.Level.LowerLimit.Enabled = true;
                    APx.SteppedFrequencySweep.Level.UpperLimit.Enabled = true;

                    APx.SteppedFrequencySweep.Level.LowerLimit.Clear();
                    APx.SteppedFrequencySweep.Level.UpperLimit.Clear();

                    if (NoUpperRefCurve)
                    {
                        APx.SteppedFrequencySweep.Level.UpperLimit.ImportData(refCurve);
                    }
                    else
                    {
                        APx.SteppedFrequencySweep.Level.UpperLimit.ImportData(refCurvePlus);
                    }

                    if (NoUpperRefCurve)
                    {
                        APx.SteppedFrequencySweep.Level.LowerLimit.ImportData(refCurve);
                    }
                    else
                    {
                        APx.SteppedFrequencySweep.Level.LowerLimit.ImportData(refCurveMinus);
                    }

                    //     APx.SteppedFrequencySweep.Level.LowerLimit.ImportData(refCurve);
                    //     APx.SteppedFrequencySweep.Level.UpperLimit.ImportData(refCurve);
                    APx.SteppedFrequencySweep.Level.LowerLimit.OffsetValues(-DUT.SweepLowerLimit);
                    APx.SteppedFrequencySweep.Level.UpperLimit.OffsetValues(DUT.SweepUpperLimit);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + " ---> " + refCurve + "\n PLEASE INFORM THE ENGINEER IN CHARGE!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    BadStatus = true;
                    DUT.hasResults = false;
                }
            }

            return BadStatus;

        } /* End of to Activate a test sequence from AP sequencer   */

        /******************************************************************************************************/
        /* Start Acoustics Response Sweep sequence for Freq Resp                                                           */
        /******************************************************************************************************/
        private void runFRTest(string FileName, string AUXCableSig)
        {
            //Thread.Sleep(50);
            if (!testSelButton9.Checked)
            {
                APx.AcousticResponse.ClearData();
            }

            try
            {
                APx.Sequence.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //Get Settled results 
            ISequenceMeasurement measurement = APx.Sequence["Signal Path1"]["Acoustic Response"];

            //Check to ensure settled results are available
            if (measurement.HasSequenceResults == true)
            {
                if (AUXCableSig == "Ring")
                {
                    if (MeasurementPassed())
                    {
                        DUT.FRTestStatus2 = "PASSED";
                    }
                    else
                    {
                        DUT.FRTestStatus2 = "FAILED";
                    }
                }
                else
                {
                    if (MeasurementPassed())
                    {
                        if (testSelButton0.Checked)
                        {
                            DUT.FRTestStatus = DUT.PFForTest0Sel[0] = "PASSED";
                        }
                        else if ((testSelButton3.Checked) && ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W")))
                        {
                            DUT.FRTestStatus = DUT.PFForTest3Sel[0] = "PASSED";
                        }
                        else
                        {
                            DUT.FRTestStatus = "PASSED";
                        }
                    }
                    else
                    {
                        if (testSelButton0.Checked)
                        {
                            DUT.FRTestStatus = DUT.PFForTest0Sel[0] = "FAILED";
                        }
                        else if ((testSelButton3.Checked) && ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W")))
                        {
                            DUT.FRTestStatus = DUT.PFForTest3Sel[0] = "FAILED";
                        }
                        else
                        {
                            DUT.FRTestStatus = "FAILED";
                        }
                    }
                }

                //Get the sequence results
                ISequenceResultCollection Results =
                  default(ISequenceResultCollection);
                Results = measurement.SequenceResults;
            }

            //           remarkBox.Text = Folder.networkDrive + Folder.testDataFolder + FileName + DUT.FRTestStatus + ".xlsx";
            //          return;

            //   APx.AcousticResponse.ExportData(Folder.networkDrive + Folder.testDataFolder + FileNameForFR + DUT.FRTestStatus + ".xlsx", NumberOfGraphPoints.GraphPointsSameAsGraph, false);
            APx.AcousticResponse.Graphs[3].Result.AsSmoothResult().ExportData(Folder.networkDrive + Folder.testDataFolder + FileName + DUT.FRTestStatus + ".xlsx", "Hz", "dBSPL");

        } /* end of Start Continuos Sweep sequence to check for FR & Polarity                                          */

        /******************************************************************************************************/
        /* Start Signal Analyzer sequence for maxSPL                                                                     */
        /******************************************************************************************************/
        private void runMaxSPLTest(int ArrayIndex, Boolean SS400Mic, string FileName)
        {
            string APTestSequence = "Signal Analyzer";

            try
            {
                APx.Sequence.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //Get Settled results
            ISequenceMeasurement measurement = APx.Sequence["Signal Path1"][APTestSequence];
            //Check to ensure settled results are available
            if (measurement.HasSequenceResults == true)
            {
                if (MeasurementPassed())
                {
                    DUT.SPLTestStatus = "PASSED";
                    if ((testSelButton3.Checked) && ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W")))
                    {
                        DUT.FRTestStatus = DUT.PFForTest3Sel[1] = "PASSED";
                    }
                    else if ((testSelButton0.Checked) || (testSelButton1.Checked) || (testSelButton2.Checked))
                    {
                        DUT.PFForTest0Sel[1] = "PASSED";
                    } //For log file

                }
                else
                {
                    DUT.SPLTestStatus = "FAILED";
                    if ((testSelButton3.Checked) && ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W")))
                    {
                        DUT.FRTestStatus = DUT.PFForTest3Sel[1] = "FAILED";
                    }
                    else if ((testSelButton0.Checked) || (testSelButton1.Checked) || (testSelButton2.Checked))
                    {
                        DUT.PFForTest0Sel[1] = "FAILED";
                    } //For log file
                }
                //Get the sequence results
                ISequenceResultCollection Results =
                  default(ISequenceResultCollection);
                Results = measurement.SequenceResults;

                /*** Get data reading from derrived result "FFT Spectrum -> x.x kHz" */
                double[] SPLReading = null;

                try
                {
                    SPLReading = Results[0].GetMeterValues();
                } //FFT Spectrum -> Specify data point (xxx kHz) 
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                DUT.SPL[ArrayIndex] = SPLReading[0];

                if ((DUT.productModel == "SS400") && (testSelButton1.Checked))
                {
                    APx.SignalAnalyzer.FFTSpectrum.ExportData(Folder.networkDrive + Folder.testDataFolder + FileName + DUT.SPLTestStatus + ".xlsx", "Hz", "V");
                }
                else
                {
                    APx.SignalAnalyzer.FFTSpectrum.ExportData(Folder.networkDrive + Folder.testDataFolder + FileName + DUT.SPLTestStatus + ".xlsx", "Hz", "dBSPL");
                }
            }
        } /* end of Start Signal Analyzer sequence for maxSPL  */

        /******************************************************************************************************/
        /* Start Signal Analyzer sequence for Output Noise Test                                               */
        /******************************************************************************************************/
        private void runNoiseTest(int ArrayIndex, string FileName)
        {
            string APTestSequence = "Signal Analyzer";

            try
            {
                APx.Sequence.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //Get Settled results
            ISequenceMeasurement measurement = APx.Sequence["Signal Path1"][APTestSequence];
            //Check to ensure settled results are available
            if (measurement.HasSequenceResults == true)
            {
                if (MeasurementPassed())
                {
                    DUT.NoiseTestStatus = "PASSED";
                }
                else
                {
                    DUT.NoiseTestStatus = "FAILED";
                }
                //Get the sequence results
                ISequenceResultCollection Results =
                  default(ISequenceResultCollection);
                Results = measurement.SequenceResults;

                /*** Get data reading from derrived result "FFT Spectrum -> Power Ave." */
                //          double[] NoiseReading = null;

                //        try { NoiseReading = Results[0].GetMeterValues(); }  
                //  catch (Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }

                double[] NoiseReading = APx.SignalMeters.LevelMeter.GetValues("dBSPL");

                DUT.SPL[ArrayIndex] = NoiseReading[0]; //Live SPL meter
                APx.SignalAnalyzer.FFTSpectrum.ExportData(Folder.networkDrive + Folder.testDataFolder + FileName + DUT.NoiseTestStatus + ".xlsx", "Hz", "dBSPL");
                //    APx.SignalAnalyzer.ExportData(      Folder.networkDrive + Folder.testDataFolder + FileName + DUT.NoiseTestStatus + ".xlsx", NumberOfGraphPoints.GraphPointsSameAsGraph, false);

            }
        } /* end of Start Signal Analyzer sequence for Noise Test  */

        /******************************************************************************************************/
        /* Start Step Freq Sweep                                                                              */
        /******************************************************************************************************/
        private void runSteppedFreqSweep(string FileName)
        {
            //Thread.Sleep(50);

            APx.SteppedFrequencySweep.ClearData();
            try
            {
                APx.Sequence.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            string APTestSequence = "Stepped Frequency Sweep";

            //Get Settled results
            ISequenceMeasurement measurement = APx.Sequence["Signal Path1"][APTestSequence];
            //Check to ensure settled results are available
            if (measurement.HasSequenceResults == true)
            {
                if (MeasurementPassed())
                {
                    DUT.FRTestStatus = DUT.PFForTest0Sel[0] = "PASSED";
                }
                else
                {
                    DUT.FRTestStatus = DUT.PFForTest0Sel[0] = "FAILED";
                }

                //Get the sequence results
                ISequenceResultCollection Results =
                  default(ISequenceResultCollection);
                Results = measurement.SequenceResults;

                try
                {
                    APx.SteppedFrequencySweep.ExportData(Folder.networkDrive + Folder.testDataFolder + FileName + DUT.FRTestStatus + ".xlsx", NumberOfGraphPoints.GraphPointsSameAsGraph, false);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }
        /* end of Step Freq Sweep sequence for LRAD-RX  

              /****************************************************************/
        /* Check for Pass/Fail test data after Sequence test was run    */
        /****************************************************************/
        private bool MeasurementPassed()
        {
            //count all of the checked signal paths and measurements in the sequence in the project file
            bool testPassed = true;
            //bool testItemPassed = true;
            string[] testResults = new string[10];
            int cnt = 0;

            foreach (ISignalPath signalPath in APx.Sequence)
            {
                if (signalPath.Checked)
                {
                    foreach (ISequenceMeasurement measurement in signalPath)
                    {
                        if (measurement.Checked)
                        {
                            foreach (ISequenceResult result in APx.Sequence[signalPath.Name][measurement.Name].SequenceResults)
                            {
                                testPassed = testPassed & (result.PassedUpperLimitCheck & result.PassedLowerLimitCheck);
                                testResults[cnt] = measurement.Name;
                                cnt++;
                            }
                        }
                    }
                }
            }

            if (APx.Sequence.Passed) return true;
            else return false;

        } /*** End of to Check for test data after Sequence test was run  ***/

        public void DisplayDebugInfo()
        {
            if (DUT.Debug == 1)
            {
                debugLabel1.Text = "Model = " + DUT.productModel + ", Network Folder = " + Folder.networkDrive + ", Test Folder = " + Folder.testDataFolder;
            }
        }

        /***********************************************************************************************************************************/
        /* Save the lastest data when any S/N was changed. 
        /***********************************************************************************************************************************/
        private void SaveData(String DetailFileName)
        {
            if ((!DUT.hasResults) || (DUT.SN[1] == "Screening") || (DUT.SN[1] == "")) return; // blank DUT.SN[1] was saved with unkown reason???

            String LogFileName = null, Serial_Numbers = null, Criteria_PassFail = null, Driver_SN = null, dBSPL_Data = null;

            if ((DUT.SN[14] == "productButton12") || (DUT.SN[14] == "productButton13")) LogFileName = "60-DEG";
            else if (DUT.SN[14] == "productButton5") LogFileName = "1000X";
            else if (DUT.SN[14] == "productButton11") LogFileName = "60-DEG w AmpPack";
            else if (DUT.SN[14] == "productButton14") LogFileName = "60-DEG wo Amp";
            else if (DUT.SN[14] == "productButton18") LogFileName = "LRAD-RX";
            else if (DUT.SN[14] == "productButton8") LogFileName = "1000XVB";
            else if (DUT.SN[14] == "productButton9") LogFileName = "1950XL";
            else if (DUT.SN[14] == "productButton10") LogFileName = "360X Manifold";
            else if (DUT.SN[14] == "productButton2") LogFileName = "450XL";
            else if (DUT.SN[14] == "productButton12") LogFileName = "60-DEG";
            else if ((DUT.SN[14] == "productButton20") || (DUT.SN[14] == "productButton21")) LogFileName = "360Xm";
            else if (DUT.SN[14] == "productButton22") LogFileName = "360XL-MID";
            else LogFileName = DUT.productModel;

            String ProgramVer_Remarks = Assembly.GetExecutingAssembly().GetName().Version.ToString() + "," + remarkBox.Text;

            String FinalResult = "PASSED";

            if ((testResult0.Text == "FAILED") || (testResult1.Text == "FAILED") || (testResult2.Text == "FAILED") || (testResult3.Text == "FAILED") ||
              (testResult4.Text == "FAILED") || (testResult5.Text == "FAILED") || (testResult6.Text == "FAILED") || (testResult7.Text == "FAILED") ||
              (testResult8.Text == "FAILED") || (testResult9.Text == "FAILED")) FinalResult = "FAILED";

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Sample's S/N, WO & General infos
            String SystemSN_General_infos = DUT.SN[1] + "," + DUT.SN[0] + "," + DateTime.Now.ToString("MM/dd/yyyy") + "," + DateTime.Now.ToString("hh:mm tt") + "," +
              DUT.WO_No + "," + FinalResult + "," + DUT.productModel + ",";

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // System's SN's
            for (int cnt = 2; cnt < 11; cnt++)
            {
                if (DUT.SN[cnt] != "")
                {
                    Serial_Numbers += DUT.SN[cnt] + ",";
                }
                if ((cnt == 5) && (DUT.productModel == "1000RX"))
                {
                    Serial_Numbers += "NA,";
                }
            }

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //Driver's S/Ns
            int NoOfDriverCnt = 1;
            for (int cnt1 = 1; cnt1 < 9; cnt1++)
            {
                if (DUT.driverSN[cnt1] != "")
                {
                    Driver_SN += DUT.driverSN[cnt1] + ",";
                    NoOfDriverCnt++;
                }
            }

            if ((DUT.SN[14] == "productButton18") && (NoOfDriverCnt == 5))
            {
                Driver_SN += "NA,NA,NA,";
            }

            if ((DUT.productModel == "360Xm 1-ST 25V-60W") || (DUT.productModel == "360Xm 1-ST 100V-60W") || (DUT.productModel == "360Xm 1-ST wo Trans"))
            {
                Driver_SN += "NA,NA,NA,";
            }
            if ((DUT.productModel == "360Xm 2-ST 100V-120W") || (DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W") || (DUT.productModel == "360Xm 2-ST wo Trans"))
            {
                Driver_SN += "NA,NA,";
            }

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////                  
            //P/F per test
            if (DUT.SN[14] == "productButton18")
            {
                Criteria_PassFail = DUT.PFForTest0Sel[0] + ",";
            } //Only FR for LRAD-RX
            else
            {
                Criteria_PassFail = DUT.PFForTest0Sel[0] + "," + DUT.PFForTest0Sel[1] + ",";
            }

            if (testSelButton1.Visible)
            {
                Criteria_PassFail += DUT.PFPerTestItem[1] + ",";
            }
            if (testSelButton2.Visible)
            {
                Criteria_PassFail += DUT.PFPerTestItem[2] + ",";
            }

            if (DUT.productModel == "360Xm 2-ST wo Trans")
            {
                Criteria_PassFail += "NA,NA,";
            }

            if (testSelButton3.Visible)
            {
                if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W") || (DUT.productModel == "360Xm 2-ST wo Trans"))
                {
                    Criteria_PassFail += DUT.PFForTest3Sel[0] + "," + DUT.PFForTest3Sel[1] + ",";
                }
                else
                {
                    Criteria_PassFail += DUT.PFPerTestItem[3] + ",";
                }
            }

            if (DUT.productModel == "360Xm 2-ST wo Trans")
            {
                Criteria_PassFail += "NA,NA,";
            }

            if (testSelButton4.Visible)
            {
                Criteria_PassFail += DUT.PFPerTestItem[4] + ",";
            }
            if (testSelButton5.Visible)
            {
                Criteria_PassFail += DUT.PFPerTestItem[5] + ",";
            }
            if (testSelButton6.Visible)
            {
                Criteria_PassFail += DUT.PFPerTestItem[6] + ",";
            }
            if (testSelButton7.Visible)
            {
                Criteria_PassFail += DUT.PFPerTestItem[7] + ",";
            }
            if (testSelButton8.Visible)
            {
                Criteria_PassFail += DUT.PFPerTestItem[8] + ",";
            }
            if (testSelButton9.Visible)
            {
                Criteria_PassFail += DUT.PFPerTestItem[9] + ",";
            }

            if (DUT.productModel == "950NXT")
            {
                Criteria_PassFail += "NA,";
            }
            if (DUT.productModel == "360Xm 4-ST 100V-240W")
            {
                Criteria_PassFail += "NA,NA,NA,NA,NA,NA,";
            }
            if ((DUT.SN[14] == "productButton4") && (AUXCableOption.Checked))
            {
                Criteria_PassFail += "NA,";
            } //If 1000V but AUX cable is NOT included
            if ((DUT.productModel == "360Xm 1-ST 25V-60W") || (DUT.productModel == "360Xm 1-ST 100V-60W") || (DUT.productModel == "360Xm 2-ST 100V-120W"))
            {
                Criteria_PassFail += "NA,NA,NA,NA,";
            }

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //Max. SPL(dB) + Noise(dB)
            if (DUT.productModel == "100X-NAVY")
            {
                dBSPL_Data = Math.Round(DUT.SPL[0], 1).ToString() + "," + Math.Round(DUT.SPL[1], 1).ToString() + "," + Math.Round(DUT.SPL[2], 1).ToString() + ",";
            }
            else if ((DUT.SN[14] == "productButton12") || (DUT.SN[14] == "productButton13") || (DUT.SN[14] == "productButton15") ||
              (DUT.SN[14] == "productButton16") || (DUT.SN[14] == "productButton10"))
            {
                if ((DUT.SN[14] == "productButton15") || (DUT.SN[14] == "productButton16") || (DUT.SN[14] == "productButton10"))
                {
                    if (DUT.productModel == "SS400")
                    {
                        dBSPL_Data = Math.Round(DUT.SPL[1], 1).ToString() + "," + Math.Round(DUT.SPL[0], 1).ToString() + ",";
                    }
                    else
                    {
                        dBSPL_Data = Math.Round(DUT.SPL[0], 1).ToString() + ",";
                    }
                }
                else
                {
                    dBSPL_Data = Math.Round(DUT.SPL[0], 1).ToString() + "," + Math.Round(DUT.SPL[2], 1).ToString() + "," + Math.Round(DUT.SPL[1], 1).ToString() + ",";
                }
            }
            //else if ((DUT.SN[14] == "productButton14") || (DUT.SN[14] == "productButton17") || (DUT.SN[14] == "productButton18") || (DUT.SN[14] == "productButton19") ||
            else if ((DUT.SN[14] == "productButton14") || (DUT.SN[14] == "productButton17") || (DUT.SN[14] == "productButton19") ||
              (DUT.SN[14] == "productButton22") || (DUT.productModel == "SS300")) //No system noise output
            {
                dBSPL_Data = Math.Round(DUT.SPL[0], 1).ToString() + ",";
            }
            else if ((DUT.SN[14] == "productButton20") || (DUT.SN[14] == "productButton21"))
            {
                if ((DUT.productModel == "360Xm 2-ST 25V-60W") || (DUT.productModel == "360Xm 2-ST 70V-60W"))
                {
                    dBSPL_Data = Math.Round(DUT.SPL[0], 1).ToString() + "," + Math.Round(DUT.SPL[3], 1).ToString() + ",";
                }
                else
                {
                    dBSPL_Data = Math.Round(DUT.SPL[0], 1).ToString() + ",NA,";
                }
            }
            else
            {
                dBSPL_Data = Math.Round(DUT.SPL[0], 1).ToString() + "," + Math.Round(DUT.SPL[1], 1).ToString() + ","; //SPL, Noise
            }

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //Combine all data strings & save                         
            try
            {
                File.SetAttributes(Folder.networkDrive + LogFileName + " Test Log" + DetailFileName + ".csv", FileAttributes.Normal);
                using (var LogFile = new StreamWriter(Folder.networkDrive + LogFileName + " Test Log" + DetailFileName + ".csv", append: true))
                {
                    LogFile.WriteLine(SystemSN_General_infos + Serial_Numbers + Driver_SN + Criteria_PassFail + dBSPL_Data + ProgramVer_Remarks); //
                    LogFile.Close();
                    File.SetAttributes(Folder.networkDrive + LogFileName + " Test Log" + DetailFileName + ".csv", FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + Folder.networkDrive + LogFileName + " Test Log" + DetailFileName + ".csv \n", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        } /*** End of SaveData */

        private void SN_Box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void TestItemSelectionChanged(object sender, EventArgs e)
        {
            setupInstruction.ForeColor = Color.Purple;
            if (testSelButton0.Checked)
            {
                setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction0");
                if (setupInstruction.Text == "")
                {
                    setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction0");
                }
            }
            else if (testSelButton1.Checked)
            {
                setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction1");
                if (setupInstruction.Text == "")
                {
                    setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction1");
                }
            }
            else if (testSelButton2.Checked)
            {
                setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction2");
                if (setupInstruction.Text == "")
                {
                    setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction2");
                }
            }
            else if (testSelButton3.Checked)
            {
                setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction3");
                if (setupInstruction.Text == "")
                {
                    setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction3");
                }
            }
            else if (testSelButton4.Checked)
            {
                setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction4");
                if (setupInstruction.Text == "")
                {
                    setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction4");
                }
            }
            else if (testSelButton5.Checked)
            {
                setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction5");
                if (setupInstruction.Text == "")
                {
                    setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction5");
                }
            }
            else if (testSelButton6.Checked)
            {
                setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction6");
                if (setupInstruction.Text == "")
                {
                    setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction6");
                }
            }
            else if (testSelButton7.Checked)
            {
                setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction7");
                if (setupInstruction.Text == "")
                {
                    setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction7");
                }
            }
            else if (testSelButton8.Checked)
            {
                setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction8");
                if (setupInstruction.Text == "")
                {
                    setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction8");
                }
            }
            else if (testSelButton9.Checked)
            {
                setupInstruction.Text = ini.IniReadValue(DUT.productModel, "SetupInstruction9");
                if (setupInstruction.Text == "")
                {
                    setupInstruction.Text = ini.IniReadValue("RefData", "SetupInstruction9");
                }
            }

        }

        private void SetupSwitchBox()
        {
            if (((DUT.productModel == "SS300") || (productButton20.Checked) || (productButton22.Checked)) && //SS300, 360Xm, 360XL-MID --> Eng test chamber
              ((testSelButton0.Checked) || (testSelButton3.Checked)))
            {
                Set_Trans_HiPWR();
            }
            else if (((DUT.productModel == "SS300") || (productButton20.Checked) || (productButton22.Checked)) && //SS300, 360Xm, 360XL-MID --> Eng test chamber
              ((testSelButton1.Checked) || (testSelButton4.Checked)))
            {
                Set_Trans_MIDPWR();
            }
            else if (((DUT.productModel == "SS300") || (productButton20.Checked) || (productButton22.Checked)) && //SS300, 360Xm, 360XL-MID --> Eng test chamber
              ((testSelButton2.Checked) || (testSelButton5.Checked)))
            {
                Set_Trans_LOWPWR();
            }
            else if ((productButton18.Checked) || (productButton19.Checked) || (productButton21.Checked) || (DUT.noBlackBoxSetup == 1)) //--> Eng test chamber LRAD-RX, 360Xm wo Trans & 2000X OR External vendor wo BlackBox setup
            {
                channelSetup(0x00, 0x00, 0x00);
            }
            else //std. production test chamber
            {
                if (DUT.ExtAmplifier == 1)
                {
                    Set_SS_UpperMic();
                    if (DUT.Debug == 1)
                    {
                        debugLabel4.Text = "External Amplifier. Upper Microphone selected";
                        debugLabel4.Update();
                    }
                }
                else if (testSelButton2.Checked)
                {
                    if ((testSelButton2.Checked) && !(productButton6.Checked) && !(productButton12.Checked) && !(productButton13.Checked) && !(DUT.productModel == "SS100") &&
                      !(DUT.productModel == "SS300") && !(productButton16.Checked) && !(productButton20.Checked))
                    {
                        MessageBox.Show("Turn down the Volume Control Knob to MINIMUM position\nWhen ready, press the OK button\n\n",
                          "Confirm", MessageBoxButtons.OK, MessageBoxIcon.Question);

                        Set_LRADX_MIC_LowerMic();

                        MessageBox.Show("Now, turn up the Volume Control Knob to MAXIMUM position\nWhen ready, press the OK button\n\n",
                          "Confirm", MessageBoxButtons.OK, MessageBoxIcon.Question);

                    }
                    if (DUT.Debug == 1)
                    {
                        debugLabel4.Text = "LRADX MIC input. Lower Microphone selected";
                        debugLabel4.Update();
                    }
                }
                else if ((DUT.productModel == "100X-NAVY") && ((testSelButton0.Checked) || (testSelButton1.Checked) || (testSelButton3.Checked) || (testSelButton4.Checked)))
                {
                    Set_100XNavy_MP3_2_LRAD_Mic();
                    if (DUT.Debug == 1)
                    {
                        debugLabel4.Text = "100X NAVY MP3-2 input. Lower Microphone selected";
                        debugLabel4.Update();
                    }
                }
                else if ((DUT.productModel == "100X-NAVY") && (testSelButton7.Checked)) //100X-NAVY IPOD
                {
                    Set_100XNavy_IPOD_LRAD_Mic();
                    if (DUT.Debug == 1)
                    {
                        debugLabel4.Text = "100X NAVY IPOD input. Lower Microphone selected";
                        debugLabel4.Update();
                    }
                }
                else if (testSelButton6.Checked)
                {
                    Set_LRADX_MP3_UpperMic();
                    if (DUT.Debug == 1)
                    {
                        debugLabel4.Text = "LRADX MP3 input. Upper Microphone selected";
                        debugLabel4.Update();
                    }
                }
                else if (productButton10.Checked)
                {
                    Set_360XManifolds_LowerMic();
                    if (DUT.Debug == 1)
                    {
                        debugLabel4.Text = "External Amplifier. Lower Microphone selected";
                        debugLabel4.Update();
                    }
                }
                else
                {
                    Set_LRADX_MP3_LowerMic();
                    if (DUT.Debug == 1)
                    {
                        debugLabel4.Text = "LRADX MP3 input. Lower Microphone selected";
                        debugLabel4.Update();
                    }
                }
            }
        }

        private void troubleShootingCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (troubleShootingCheckBox.Checked)
            {
                SN_Box1.Text = "";
                recordSampleInfo(true);
            }
            else
            {
                recordSampleInfo(false);
            }

            resetreportFormat();
            setupTestSelection();

        }

        private void AUXCableOption_CheckedChanged(object sender, EventArgs e)
        {
            if (AUXCableOption.Checked)
            {
                testSelButton9.Visible = false;
                testResult9.Visible = false;
            }
            else
            {
                testSelButton9.Visible = true;
                testResult9.Visible = true;
            }
        }

        /*** Kevin's code starts here ***/
        private string GetPathFromProductModelStr()
        {
            const string PREFIX = "Z:\\Kevin Pham\\"; // Folder.networkDrive
            const string SUFFIX = " Test Log (Post-Test).csv";
            string infix;
            switch (DUT.productModel)
            {
                case "450XL Extended Test":
                    infix = "450XL";
                    break;
                case "1000X-20FT":
                case "1000X-30FT":
                case "1000X-100FT":
                    infix = "1000X";
                    break;
                case "1000XVB-20FT":
                case "1000XVB-30FT":
                case "1000XVB-100FT":
                    infix = "1000XVB";
                    break;
                case "1950XL-35FT":
                case "1950XL-66FT":
                case "1950XL-100FT":
                    infix = "1950XL";
                    break;
                case "360XL Manifold":
                    infix = "360X Manifold";
                    break;
                case "DS60-X w AmpPack":
                case "DS60-XL w AmpPack":
                    infix = "60-DEG w AmpPack";
                    break;
                case "DS60-70V-60W":
                case "DS60-100V-60W":
                case "DS60-25V-80W":
                case "DS60-70V-80W":
                case "DS60-100V-80W":
                case "DS60-70V-160W":
                case "DS60-100V-160W":
                    infix = "60-DEG";
                    break;
                case "DS60-X":
                case "DS60-XL":
                    infix = "60-DEG wo Amp";
                    break;
                case "500RX":
                case "950RXL":
                case "1000RX":
                case "950NXT":
                    infix = "LRAD-RX";
                    break;
                case "360Xm 1-ST 25V-60W":
                case "360Xm 1-ST 100V-60W":
                case "360Xm 2-ST 25V-60W":
                case "360Xm 2-ST 100V-120W":
                case "360Xm 2-ST 70V-60W":
                case "360Xm 4-ST 100V-240W":
                case "360Xm 1-ST wo Trans":
                case "360Xm 2-ST wo Trans":
                    infix = "360Xm";
                    break;
                case "360XL-MID 1-ST":
                case "360XL-MID 2-ST":
                    infix = "360XL-MID";
                    break;
                default:
                    infix = DUT.productModel;
                    break;
            }
            return PREFIX + infix + SUFFIX;
        }

        private List<List<string>> GetAllUnitSnTests(string path, string unitSn)
        {
            List<List<string>> allTestsFromSn = new List<List<string>>();
            using (StreamReader reader = new StreamReader(path))
            {
                string headerLine = reader.ReadLine();
                string line;
                while (!reader.EndOfStream)
                {
                    line = reader.ReadLine();
                    var values = line.Split(';');
                    List<String> row = values[0].Split(',').ToList();
                    if (row[0].Trim() == unitSn)
                    {
                        int headerRowSize = File.ReadLines(path).First().Split(',').Length;
                        int selectedRowSize = row.Count;
                        int elementsNeeded = headerRowSize - selectedRowSize;
                        if (elementsNeeded > 0)
                        {
                            for (int i = 0; i < elementsNeeded; i++)
                            {
                                row.Add("");
                            }
                        }
                        allTestsFromSn.Add(new List<string>(row));
                    }
                }
            }
            return allTestsFromSn;
        }

        Model GetModelFromProductModelStr(List<string> latestTest)
        {
            Model model;
            switch (DUT.productModel)
            {
                case "100X":
                    model = new Model100X(latestTest);
                    break;
                case "100X-NAVY-V1":
                    model = new Model100XNavyV01(latestTest);
                    break;
                case "100X-NAVY":
                    model = new Model100XNavy(latestTest);
                    break;
                case "300Xi":      
                case "300XRA":
                case "300XRA-260W":
                    model = new Model300X(latestTest);
                    break;
                case "450XL":
                case "450XL Extended Test":
                case "450XL-RA":
                    model = new Model450Xl(latestTest);
                    break;
                case "500X":
                    model = new Model500X(latestTest);
                    break;
                case "500X-RE":
                    model = new Model500Xre(latestTest);
                    break;
                case "1000":
                    model = new Model1000V(latestTest);
                    break;
                case "1000X-20FT":
                case "1000X-30FT":
                case "1000X-100FT":
                    model = new Model1000X(latestTest);
                    break;
                case "1000X2U":
                    model = new Model1000X2U(latestTest);
                    break;
                case "1000Xi":
                    model = new Model1000Xi(latestTest);
                    break;
                case "1000XVB-20FT":
                case "1000XVB-30FT":
                case "1000XVB-100FT":
                    model = new Model1000Xvb(latestTest);
                    break;
                case "1950XL-35FT":
                case "1950XL-66FT":
                case "1950XL-100FT":
                    model = new Model1950Xl(latestTest);
                    break;
                case "360X Manifold":
                case "360XL Manifold":
                    model = new Model360XManifolds(latestTest);
                    break;
                case "DS60-X w AmpPack":
                case "DS60-XL w AmpPack":
                    model = new ModelDs60WAmpPack(latestTest);
                    break;
                case "DS60-70V-60W":
                case "DS60-100V-60W":
                case "DS60-25V-80W":
                case "DS60-70V-80W":
                case "DS60-100V-80W":
                case "DS60-70V-160W":
                case "DS60-100V-160W":
                    model = new ModelDs60X(latestTest);
                    break;
                case "DS60-X":
                case "DS60-XL":
                    model = new ModelDs60HornsOnly(latestTest);
                    break;
                case "SS100":
                    model = new ModelSS100(latestTest);
                    break;
                case "SS300":
                    model = new ModelSS300(latestTest);
                    break;
                case "SS400":
                    model = new ModelSS400(latestTest);
                    break;
                case "SSX":
                case "SSX60":
                    model = new ModelSsx(latestTest);
                    break;
                case "SSX wo Trans":
                case "SSX60 wo Trans":
                    model = new ModelSsxWoTrans(latestTest);
                    break;
                case "500RX":
                case "950RXL":
                case "1000RX":
                case "950NXT":
                    model = new ModelLradRx(latestTest);
                    break;
                case "2000X":
                    model = new Model2000X(latestTest);
                    break;
                case "360Xm 1-ST 25V-60W":
                case "360Xm 1-ST 100V-60W":
                case "360Xm 2-ST 25V-60W":
                case "360Xm 2-ST 100V-120W":
                case "360Xm 2-ST 70V-60W":
                case "360Xm 4-ST 100V-240W":
                case "360Xm 1-ST wo Trans":
                case "360Xm 2-ST wo Trans":
                    model = new Model360Xm(latestTest);
                    break;
                case "360XL-MID 1-ST":
                case "360XL-MID 2-ST":
                    model = new Model300XlMid(latestTest);
                    break;
                default:
                    throw new ArgumentException("productModel");
            }
            return model;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            const string ERR_1 = "Please enter a System SN!";
            const string ERR_2 = "No tests found!";
            const string WARN_1 = "Multiple tests found. Latest test shown.";

            string path = GetPathFromProductModelStr();
            string unitSn = SN_Box1.Text;
            if (unitSn.Trim() == "")
            {
                MessageBox.Show(ERR_1, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            List<List<string>> tests = GetAllUnitSnTests(path, unitSn);
            switch(tests.Count)
            {
                case 0:
                    MessageBox.Show(ERR_2, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                case 1:
                    break;
                default:
                    MessageBox.Show(WARN_1, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
            List<string> latestTest = tests.Last();
            Model model = GetModelFromProductModelStr(latestTest);
        }
    }
}