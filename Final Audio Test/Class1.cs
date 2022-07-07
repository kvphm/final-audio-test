using AudioPrecision.API;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

/************************************************************************************************/
/*  Custom Product class                                                                        */
/************************************************************************************************/

namespace Final_Audio_Test
{
    public class productionDrive
    {
        public string setupPicsFolder = null;
        public string refFileFolder = null;
        public string networkDrive = null;
        public string networkDriveForSS = null;
        public string networkDriveForBenchTest = null;
        public string testDataFolder = null;  //Raw data for eack model name
    }
    public class Product
    {
        //Measured data
        public double[] SPL = new double[5];  //0=Max SPL, 1=Noise Output, 2=IPOD Max SPL    
        public string[] SN = new string[16];     //hold SNs for log file. SN[0]=Oper Ini. SN[14]=string(productButtonID) --> To hold data for Product Series; SN[15]=string(productModel) --> To hold data for Product Model
        public string[] driverSN = new string[16];     
        public string[] PFForTest0Sel = new string[2];     //Hold Pass/Fail Status for Sensitivity test --> PFForTest0Sel[0]=Sensititivity, PFForTest0Sel[1]=Max SPL       
        public string[] PFForTest3Sel = new string[2];      //Same as above but for 360Xm 2-ST w Trans
        public string[] PFPerTestItem = new string[16];     //Hold Pass/Fail Status per test item --> 1->9 is per Test Selection, 10-> Up are for extra test items are NOT in Test Selection
                                                            //PFperTestItem[0] is not used --> Refer to PFForTest0Sel[x]

        public string productModel;
        public string FileNameExt;
        public string WO_No;

        public string FRTestStatus;
        public string FRTestStatus2;  //Only for 1000V AUX cable test --> Ring Signal
        public string SPLTestStatus;      
        public string NoiseTestStatus;

        public byte portCState;

        public int freqStart;     //sweep start Freq
        public int freqStop;
        public int freqSPL;       //Max SPL Freq & for SS400 built-In recording mic
        public int autoShutdown;
        public int APautoShutdown;
        public int noBlackBoxSetup;  //Uses for outside vendor --> No blackbox selecting, test cables are directly connect to AP and DUT
        public int Debug;
        public int ExtAmplifier;  //1=Products w External amp

        public double WindowBegin;
        public double WindowEnd;

        public double stimulusFR;
        public double stimulusSPL;

        public double SweepUpperLimit;
        public double SweepLowerLimit;
        public double SPLLowerLimit; //dBSPL
        public double SPLUpperLimit; //           

        public double YAxisMax;
        public double YAxisMin;
        public double XAxisMax;
        public double XAxisMin;
        public double YAxisMaxSPL;
        public double YAxisMinSPL;

        public string graphTitle;
        public string SPLGraphTitle;
        public string SS400MicInputGraphTitle;
        public string ReportFormPN;

        public string curveSmooth;

        //Ref SPL levels
        public double RefSPL; //SPL Ref
        public double RefIPODSPL; //SPL Ref for IPOD input in 100X-NAVY
        public double RefSS400MicInput; //Recording Mic SPL Ref

        //Trouble shooting mode & Load Data Flags        
        public Boolean loadDataFlag = false;
        public Boolean hasResults = false;

        public Product()
        {
            resetVariablesToDefault();
        }

        public void resetVariablesToDefault()
        {
            for (int cnt = 0; cnt < 16; cnt++)
            {
                SN[cnt] = "";
                driverSN[cnt] = "";
                PFPerTestItem[cnt] = "NT";
            }

            for (int cnt = 0; cnt < 5; cnt++) { SPL[cnt] = 0.0; }

            PFForTest0Sel[0] = "NT";
            PFForTest0Sel[1] = "NT";

            PFForTest3Sel[0] = "NT";
            PFForTest3Sel[1] = "NT";

            FRTestStatus = "NT";
            FRTestStatus2 = "NT";
            SPLTestStatus = "NT";
            NoiseTestStatus = "NT";                       
            hasResults = false;
        }
    }

}

/************************************************************************************************/
/*  Custom Ini class                                                                            */
/************************************************************************************************/
namespace Ini
{
    /// <summary>
    /// Create a New INI file to store or load data
    /// </summary>
    public class IniFile
    {
        public string path;

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section,
            string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section,
                 string key, string def, StringBuilder retVal,
            int size, string filePath);

        /// <summary>
        /// INIFile Constructor.
        /// </summary>
        /// <PARAM name="INIPath"></PARAM>
        public IniFile(string INIPath)
        {
            path = INIPath;
        }
        /// <summary>
        /// Write Data to the INI File
        /// </summary>
        /// <PARAM name="Section"></PARAM>
        /// Section name
        /// <PARAM name="Key"></PARAM>
        /// Key Name
        /// <PARAM name="Value"></PARAM>
        /// Value Name
        public void IniWriteValue(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, this.path);
        }

        /// <summary>
        /// Read Data Value From the Ini File
        /// </summary>
        /// <PARAM name="Section"></PARAM>
        /// <PARAM name="Key"></PARAM>
        /// <PARAM name="Path"></PARAM>
        /// <returns></returns>
        public string IniReadValue(string Section, string Key)
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp,
                                            255, this.path);
            return temp.ToString();

        }
    }
}
/**********************************************************************************************/
