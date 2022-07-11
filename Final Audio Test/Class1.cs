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

    // Kevin's Code:
    class Field
    {
        private static string VALUE_PASSED = "PASSED";
        private static string VALUE_FAILED = "FAILED";
        private static string VALUE_NT = "NT";

        private int _index;
        private string _value;

        public int Index
        {
            get 
            { 
                return _index; 
            } 
            set 
            { 
                this._index = value; 
            } 
        }

        public string Value
        {
            get
            { 
                return _value; 
            }
            set 
            { 
                this._value = value; 
            }
        }

        public Field(int index, string value)
        {
            this._index = index;
            this._value = value;
        }

        public static string TwoPFEval(string data1, string data2)
        {
            if (data1.Equals(VALUE_PASSED) && data2.Equals(VALUE_PASSED))
            {
                return VALUE_PASSED;
            }
            else if (data1.Equals(VALUE_PASSED) || data2.Equals(VALUE_PASSED))
            {
                return VALUE_FAILED;
            }
            else
            {
                return VALUE_NT;
            }
        }
    }

    abstract class Model
    {
        protected const int COL_UNIT_SN   = 0;
        protected const int COL_OPERATOR  = 1;
        protected const int COL_DATE      = 2;
        protected const int COL_TIME      = 3;
        protected const int COL_WO_NO     = 4;
        protected const int COL_SYSTEM_PF = 5;

        protected const int INDEX_OPER_INI    = 0;
        protected const int INDEX_WO_NO       = 1;
        protected const int INDEX_DATE        = 2;
        protected const int INDEX_TIME        = 3;
        
        protected const int INDEX_SYSTEM      = 0;

        protected const int INDEX_DRIVER_SN_1 = 0;
        protected const int INDEX_DRIVER_SN_2 = 1;
        protected const int INDEX_DRIVER_SN_3 = 2;
        protected const int INDEX_DRIVER_SN_4 = 3;
        protected const int INDEX_DRIVER_SN_5 = 4;
        protected const int INDEX_DRIVER_SN_6 = 5;
        protected const int INDEX_DRIVER_SN_7 = 6;
        protected const int INDEX_DRIVER_SN_8 = 7;

        protected const int INDEX_REMARKS     = 4;

        protected List<Field> general;
        protected List<Field> serialNo;
        protected List<Field> driverSns;
        protected List<Field> subTests;

        protected Model(string[] data)
        {
            general = new List<Field>();
            serialNo = new List<Field>();
            driverSns = new List<Field>();
            subTests = new List<Field>();

            general.Add(new Field(INDEX_OPER_INI, data[COL_OPERATOR]));
            general.Add(new Field(INDEX_WO_NO, data[COL_WO_NO]));
            general.Add(new Field(INDEX_DATE, data[COL_DATE]));
            general.Add(new Field(INDEX_TIME, data[COL_TIME]));
            
            serialNo.Add(new Field(INDEX_SYSTEM, data[COL_UNIT_SN]));
        }
    }

    class Model100X : Model
    {
        private const int COL_MP3_PLAYER_SN   = 7;
        private const int COL_ELEC_SN         = 8;
        private const int COL_MIC_SN          = 9;
        private const int COL_BATT_SN         = 10;

        private const int COL_DRIVER_SN       = 11;

        private const int COL_MP3_SENSITIVITY = 12;
        private const int COL_MAX_SPL         = 13;
        private const int COL_VOLUME_FUNCTION = 14;
        private const int COL_MIC_SENSITIVITY = 15;
        private const int COL_OUTPUT_NOISE    = 16;

        private const int COL_REMARKS         = 20;

        private const int INDEX_MP3_PLAYER                     = 1;
        private const int INDEX_ELECTRONICS                    = 2;
        private const int INDEX_MICROPHONE                     = 3;
        private const int INDEX_BATTERY                        = 4;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL  = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION   = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY   = 2;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT = 6;

        Model100X(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_MP3_PLAYER, data[COL_MP3_PLAYER_SN]));
            serialNo.Add(new Field(INDEX_ELECTRONICS, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));
            serialNo.Add(new Field(INDEX_BATTERY, data[COL_BATT_SN]));
            
            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOLUME_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_OUTPUT_NOISE]));
        }
    }

    class Model100XNavyV01 : Model
    {
        private const int COL_MP3_PLAYER_SN   = 7;
        private const int COL_ELEC_SN         = 8;
        private const int COL_MIC_SN          = 9;
        private const int COL_BATT_SN         = 10;

        private const int COL_DRIVER_SN       = 11;

        private const int COL_MP3_SENSITIVITY = 12;
        private const int COL_MAX_SPL         = 13;
        private const int COL_VOLUME_FUNCTION = 14;
        private const int COL_MIC_SENSITIVITY = 15;
        private const int COL_OUTPUT_NOISE    = 16;

        private const int COL_REMARKS         = 20;

        private const int INDEX_MP3_PLAYER                     = 1;
        private const int INDEX_ELECTRONICS                    = 2;
        private const int INDEX_MICROPHONE                     = 3;
        private const int INDEX_BATTERY                        = 4;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL  = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION   = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY   = 2;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT = 6;

        Model100XNavyV01(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_MP3_PLAYER, data[COL_MP3_PLAYER_SN]));
            serialNo.Add(new Field(INDEX_ELECTRONICS, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));
            serialNo.Add(new Field(INDEX_BATTERY, data[COL_BATT_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOLUME_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT,data[COL_OUTPUT_NOISE]));
        }
    }

    class Model100XNavy : Model
    {
        private const int COL_ELEC_SN          = 7;
        private const int COL_MIC_SN           = 8;
        private const int COL_BATT_SN          = 9;

        private const int COL_DRIVER_SN        = 10;

        private const int COL_MP32_SENSITIVITY = 11;
        private const int COL_MAX_SPL          = 12;
        private const int COL_VOL_FUNCTION     = 13;
        private const int COL_MIC_SENSITIVITY  = 14;
        private const int COL_PWR_ON_VB_ON     = 15;
        private const int COL_PWR_OFF_VB_OFF   = 16;
        private const int COL_MP31_SENSITIVITY = 17;
        private const int COL_OUTPUT_NOISE     = 18;
        private const int COL_IPOD_INPUT       = 19;
        private const int COL_MAX_SPL_DB       = 20;

        private const int COL_REMARKS          = 24;

        private const int INDEX_ELECTRONICS                         = 2;
        private const int INDEX_MICROPHONE                          = 3;
        private const int INDEX_BATTERY                             = 4;

        private const int INDEX_MP3_INPUT2_SENSITIVITY_MAX_SPL      = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION        = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY        = 2;
        private const int INDEX_AUDIO_EFFECT_WITH_PWR_SW_ON_VB_ON   = 3;
        private const int INDEX_AUDIO_EFFECT_WITH_PWR_SW_OFF_VB_OFF = 4;
        private const int INDEX_MP3_INPUT1_SENSITIVITY              = 5;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT      = 6;
        private const int INDEX_IPOD_INPUT_SENSITIVITY_MAX_SPL      = 7;

        Model100XNavy(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_ELECTRONICS, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));
            serialNo.Add(new Field(INDEX_BATTERY, data[COL_BATT_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT2_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP32_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_WITH_PWR_SW_ON_VB_ON, data[COL_PWR_ON_VB_ON]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_WITH_PWR_SW_OFF_VB_OFF, data[COL_PWR_OFF_VB_OFF]));
            subTests.Add(new Field(INDEX_MP3_INPUT1_SENSITIVITY, data[COL_MP31_SENSITIVITY]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_OUTPUT_NOISE]));
            subTests.Add(new Field(INDEX_IPOD_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_IPOD_INPUT], data[COL_MAX_SPL_DB])));
        }
    }

    class Model300Xi : Model
    {
        private const int COL_CONTROL_UNIT_SN = 7;
        private const int COL_ELEC_SN         = 8;
        private const int COL_MIC_SN          = 9;

        private const int COL_DRIVER1_SN      = 10;
        private const int COL_DRIVER2_SN      = 11;

        private const int COL_MP3_SENSITIVITY = 12;
        private const int COL_MAX_SPL         = 13;
        private const int COL_VOL_FUNCTION    = 14;
        private const int COL_MIC_SENSITIVITY = 15;
        private const int COL_WIDE_ON         = 16;
        private const int COL_NARROW_OFF      = 17;
        private const int COL_NARROW_ON       = 18;
        private const int COL_OUTPUT_NOISE    = 19;

        private const int COL_REMARKS         = 23;

        private const int INDEX_CONTROL_UNIT                       = 1;
        private const int INDEX_ELECTRONICS                        = 2;
        private const int INDEX_MICROPHONE                         = 3;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL      = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION       = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY       = 2;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON    = 3;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF = 4;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON  = 5;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT     = 6;

        Model300Xi(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_CONTROL_UNIT, data[COL_CONTROL_UNIT_SN]));
            serialNo.Add(new Field(INDEX_ELECTRONICS, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON, data[COL_WIDE_ON]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF, data[COL_NARROW_OFF]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON, data[COL_NARROW_ON]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_OUTPUT_NOISE]));
        }
    }

    class Model300XRa : Model
    {
        private const int COL_CONTROL_UNIT_SN = 7;
        private const int COL_ELEC_SN         = 8;
        private const int COL_MIC_SN          = 9;

        private const int COL_DRIVER1_SN      = 10;
        private const int COL_DRIVER2_SN      = 11;

        private const int COL_MP3_SENSITIVITY = 12;
        private const int COL_MAX_SPL         = 13;
        private const int COL_VOL_FUNCTION    = 14;
        private const int COL_MIC_SENSITIVITY = 15;
        private const int COL_WIDE_ON         = 16;
        private const int COL_NARROW_OFF      = 17;
        private const int COL_NARROW_ON       = 18;
        private const int COL_OUTPUT_NOISE    = 19;

        private const int COL_REMARKS         = 23;

        private const int INDEX_CONTROL_UNIT                       = 1;
        private const int INDEX_ELECTRONICS                        = 2;
        private const int INDEX_MICROPHONE                         = 3;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL      = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION       = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY       = 2;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON    = 3;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF = 4;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON  = 5;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT     = 6;

        Model300XRa(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_CONTROL_UNIT, data[COL_CONTROL_UNIT_SN]));
            serialNo.Add(new Field(INDEX_ELECTRONICS, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON, data[COL_WIDE_ON]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF, data[COL_NARROW_OFF]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON, data[COL_NARROW_ON]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_OUTPUT_NOISE]));
        }
    }

    class Model300Xra260W : Model
    {
        private const int COL_CONTROL_UNIT_SN = 7;
        private const int COL_ELEC_SN         = 8;
        private const int COL_MIC_SN          = 9;

        private const int COL_DRIVER1_SN      = 10;
        private const int COL_DRIVER2_SN      = 11;

        private const int COL_MP3_SENSITIVITY = 12;
        private const int COL_MAX_SPL         = 13;
        private const int COL_VOL_FUNCTION    = 14;
        private const int COL_MIC_SENSITIVITY = 15;
        private const int COL_WIDE_ON         = 16;
        private const int COL_NARROW_OFF      = 17;
        private const int COL_NARROW_ON       = 18;
        private const int COL_OUTPUT_NOISE    = 19;

        private const int COL_REMARKS         = 23;

        private const int INDEX_CONTROL_UNIT                       = 1;
        private const int INDEX_ELECTRONICS                        = 2;
        private const int INDEX_MICROPHONE                         = 3;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL      = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION       = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY       = 2;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON    = 3;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF = 4;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON  = 5;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT     = 6;

        Model300Xra260W(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_CONTROL_UNIT, data[COL_CONTROL_UNIT_SN]));
            serialNo.Add(new Field(INDEX_ELECTRONICS, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON, data[COL_WIDE_ON]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF, data[COL_NARROW_OFF]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON, data[COL_NARROW_ON]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_OUTPUT_NOISE]));
        }
    }

    class Model450Xl : Model
    {
        private const int COL_CONTROL_UNIT_SN = 7;
        private const int COL_ELEC_SN         = 8;
        private const int COL_MIC_SN          = 9;

        private const int COL_DRIVER1_SN      = 10;
        private const int COL_DRIVER2_SN      = 11;

        private const int COL_MP3_SENSITIVITY = 12;
        private const int COL_MAX_SPL         = 13;
        private const int COL_VOL_FUNCTION    = 14;
        private const int COL_MIC_SENSITIVITY = 15;
        private const int COL_WIDE_ON         = 16;
        private const int COL_NARROW_OFF      = 17;
        private const int COL_NARROW_ON       = 18;
        private const int COL_OUTPUT_NOISE    = 19;

        private const int COL_REMARKS         = 23;

        private const int INDEX_CONTROL_UNIT                       = 1;
        private const int INDEX_ELECTRONICS                        = 2;
        private const int INDEX_MICROPHONE                         = 3;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL      = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION       = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY       = 2;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON    = 3;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF = 4;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON  = 5;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT     = 6;

        Model450Xl(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_CONTROL_UNIT, data[COL_CONTROL_UNIT_SN]));
            serialNo.Add(new Field(INDEX_ELECTRONICS, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON, data[COL_WIDE_ON]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF, data[COL_NARROW_OFF]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON, data[COL_NARROW_ON]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_OUTPUT_NOISE]));
        }
    }

    class Model450XlExtendedTest : Model
    {
        private const int COL_CONTROL_UNIT_SN = 7;
        private const int COL_ELEC_SN         = 8;
        private const int COL_MIC_SN          = 9;

        private const int COL_DRIVER1_SN      = 10;
        private const int COL_DRIVER2_SN      = 11;

        private const int COL_MP3_SENSITIVITY = 12;
        private const int COL_MAX_SPL         = 13;
        private const int COL_VOL_FUNCTION    = 14;
        private const int COL_MIC_SENSITIVITY = 15;
        private const int COL_WIDE_ON         = 16;
        private const int COL_NARROW_OFF      = 17;
        private const int COL_NARROW_ON       = 18;
        private const int COL_OUTPUT_NOISE    = 19;

        private const int COL_REMARKS         = 23;

        private const int INDEX_CONTROL_UNIT                       = 1;
        private const int INDEX_ELECTRONICS                        = 2;
        private const int INDEX_MICROPHONE                         = 3;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL      = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION       = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY       = 2;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON    = 3;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF = 4;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON  = 5;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT     = 6;

        Model450XlExtendedTest(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_CONTROL_UNIT, data[COL_CONTROL_UNIT_SN]));
            serialNo.Add(new Field(INDEX_ELECTRONICS, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON, data[COL_WIDE_ON]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF, data[COL_NARROW_OFF]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON, data[COL_NARROW_ON]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_OUTPUT_NOISE]));
        }
    }

    class Model450XlRa : Model
    {
        private const int COL_CONTROL_UNIT_SN = 7;
        private const int COL_ELEC_SN         = 8;
        private const int COL_MIC_SN          = 9;

        private const int COL_DRIVER1_SN      = 10;
        private const int COL_DRIVER2_SN      = 11;

        private const int COL_MP3_SENSITIVITY = 12;
        private const int COL_MAX_SPL         = 13;
        private const int COL_VOL_FUNCTION    = 14;
        private const int COL_MIC_SENSITIVITY = 15;
        private const int COL_WIDE_ON         = 16;
        private const int COL_NARROW_OFF      = 17;
        private const int COL_NARROW_ON       = 18;
        private const int COL_OUTPUT_NOISE    = 19;

        private const int COL_REMARKS         = 23;

        private const int INDEX_CONTROL_UNIT                       = 1;
        private const int INDEX_ELECTRONICS                        = 2;
        private const int INDEX_MICROPHONE                         = 3;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL      = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION       = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY       = 2;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON    = 3;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF = 4;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON  = 5;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT     = 6;

        Model450XlRa(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_CONTROL_UNIT, data[COL_CONTROL_UNIT_SN]));
            serialNo.Add(new Field(INDEX_ELECTRONICS, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON, data[COL_WIDE_ON]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF, data[COL_NARROW_OFF]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON, data[COL_NARROW_ON]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_OUTPUT_NOISE]));
        }
    }

    class Model500X : Model
    {
        private const int COL_MP3_PLAYER_SN   = 7;
        private const int COL_ELEC_SN         = 8;
        private const int COL_MIC_SN          = 9;

        private const int COL_DRIVER1_SN      = 10;
        private const int COL_DRIVER2_SN      = 11;
        private const int COL_DRIVER3_SN      = 12;
        private const int COL_DRIVER4_SN      = 13;

        private const int COL_MP3_SENSITIVITY = 14;
        private const int COL_MAX_SPL         = 15;
        private const int COL_VOL_FUNCTION    = 16;
        private const int COL_MIC_SENSITIVITY = 17;
        private const int COL_WIDE_ON         = 18;
        private const int COL_NARROW_OFF      = 19;
        private const int COL_NARROW_ON       = 20;
        private const int COL_OUTPUT_NOISE    = 21;

        private const int COL_REMARKS         = 25;

        private const int INDEX_MP3_PLAYER                         = 1;
        private const int INDEX_ELECTRONICS                        = 2;
        private const int INDEX_MICROPHONE                         = 3;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL      = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION       = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY       = 2;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON    = 3;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF = 4;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON  = 5;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT     = 6;

        Model500X(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_MP3_PLAYER, data[COL_MP3_PLAYER_SN]));
            serialNo.Add(new Field(INDEX_ELECTRONICS, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_3, data[COL_DRIVER3_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_4, data[COL_DRIVER4_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON, data[COL_WIDE_ON]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF, data[COL_NARROW_OFF]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON, data[COL_NARROW_ON]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_OUTPUT_NOISE]));
        }
    }

    class Model500Xre : Model
    {
        private const int COL_CONTROL_UNIT_SN = 7;
        private const int COL_ELEC_SN         = 8;
        private const int COL_MIC_SN          = 9;

        private const int COL_DRIVER1_SN      = 10;
        private const int COL_DRIVER2_SN      = 11;
        private const int COL_DRIVER3_SN      = 12;
        private const int COL_DRIVER4_SN      = 13;

        private const int COL_MP3_SENSITIVITY = 14;
        private const int COL_MAX_SPL         = 15;
        private const int COL_VOL_FUNCTION    = 16;
        private const int COL_MIC_SENSITIVITY = 17;
        private const int COL_WIDE_ON         = 18;
        private const int COL_NARROW_OFF      = 19;
        private const int COL_NARROW_ON       = 20;
        private const int COL_OUTPUT_NOISE    = 21;

        private const int COL_REMARKS         = 25;

        private const int INDEX_CONTROL_UNIT                       = 1;
        private const int INDEX_ELECTRONICS                        = 2;
        private const int INDEX_MICROPHONE                         = 3;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL      = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION       = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY       = 2;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON    = 3;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF = 4;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON  = 5;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT     = 6;

        Model500Xre(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_CONTROL_UNIT, data[COL_CONTROL_UNIT_SN]));
            serialNo.Add(new Field(INDEX_ELECTRONICS, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_3, data[COL_DRIVER3_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_4, data[COL_DRIVER4_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON, data[COL_WIDE_ON]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF, data[COL_NARROW_OFF]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON, data[COL_NARROW_ON]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_OUTPUT_NOISE]));
        }
    }

    class Model1000V : Model
    {
        private const int COL_X_MP3_AL        = 7;
        private const int COL_1000_AMP        = 8;
        private const int COL_X_MIC_AL        = 9;
        private const int COL_1000_G_SYS      = 10;
        private const int COL_1000_AC_PWR     = 11;
        private const int COL_PHRASSL_TR_2    = 12;

        private const int COL_DRIVER1_SN      = 13;
        private const int COL_DRIVER2_SN      = 14;
        private const int COL_DRIVER3_SN      = 15;
        private const int COL_DRIVER4_SN      = 16;

        private const int COL_MP3_SENSITIVITY = 17;
        private const int COL_MAX_SPL         = 18;
        private const int COL_VOL_FUNCTION    = 19;
        private const int COL_MIC_SENSITIVITY = 20;
        private const int COL_WIDE_ON         = 21;
        private const int COL_NARROW_OFF      = 22;
        private const int COL_NARROW_ON       = 23;
        private const int COL_OUTPUT_NOISE    = 24;
        private const int COL_LIMITED_KEY_SW  = 25;
        private const int COL_SELF_TEST       = 26;
        private const int COL_AUX_CABLE       = 27;

        private const int COL_REMARKS         = 31;

        private const int INDEX_X_MP3_AL                           = 1;
        private const int INDEX_1000_AMP                           = 2;
        private const int INDEX_X_MIC_AL                           = 3;
        private const int INDEX_1000_G_SYS                         = 4;
        private const int INDEX_1000_AC_PWR                        = 5;
        private const int INDEX_PHRASLTR_2                         = 6;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL      = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION       = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY       = 2;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON    = 3;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF = 4;
        private const int INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON  = 5;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT     = 6;
        private const int INDEX_LIMITED_KEY_SWITCH_FUNCTION        = 7;
        private const int INDEX_SELF_TEST                          = 8;
        private const int INDEX_AUX_CABLE                          = 9;

        Model1000V(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_X_MP3_AL, data[COL_X_MP3_AL]));
            serialNo.Add(new Field(INDEX_1000_AMP, data[COL_1000_AMP]));
            serialNo.Add(new Field(INDEX_X_MIC_AL, data[COL_X_MIC_AL]));
            serialNo.Add(new Field(INDEX_1000_G_SYS, data[COL_1000_G_SYS]));
            serialNo.Add(new Field(INDEX_1000_AC_PWR, data[COL_1000_AC_PWR]));
            serialNo.Add(new Field(INDEX_PHRASLTR_2, data[COL_PHRASSL_TR_2]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_3, data[COL_DRIVER3_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_4, data[COL_DRIVER4_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_WIDE_VB_ON, data[COL_WIDE_ON]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_OFF, data[COL_NARROW_OFF]));
            subTests.Add(new Field(INDEX_AUDIO_EFFECT_W_SOUND_NARROW_VB_ON, data[COL_NARROW_ON]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_OUTPUT_NOISE]));
            subTests.Add(new Field(INDEX_LIMITED_KEY_SWITCH_FUNCTION, data[COL_LIMITED_KEY_SW]));
            subTests.Add(new Field(INDEX_SELF_TEST, data[COL_SELF_TEST]));
            subTests.Add(new Field(INDEX_AUX_CABLE, data[COL_AUX_CABLE]));
        }
    }

    class Model1000X20AudioCable : Model
    {
        private const int COL_MP3_PLAYER_SN                  = 7;
        private const int COL_ELEC_SN                        = 8;
        private const int COL_MIC_SN                         = 9;

        private const int COL_DRIVER1_SN                     = 10;
        private const int COL_DRIVER2_SN                     = 11;
        private const int COL_DRIVER3_SN                     = 12;
        private const int COL_DRIVER4_SN                     = 13;
        private const int COL_DRIVER5_SN                     = 14;
        private const int COL_DRIVER6_SN                     = 15;
        private const int COL_DRIVER7_SN                     = 16;

        private const int COL_MP3_SENSITIVITY                = 17;
        private const int COL_MAX_SPL                        = 18;
        private const int COL_VOL_FUNCTION                   = 19;               
        private const int COL_MIC_SENSITIVITY                = 20;
        private const int COL_AUDIO_OUTPUT_SW_OFF            = 21;
        private const int COL_SYSTEM_BACKGROUND_NOISE_OUTPUT = 22;

        private const int COL_REMARKS                        = 26;

        private const int INDEX_MP3_PLAYER                                   = 1;
        private const int INDEX_AMP_PACK                                     = 2;
        private const int INDEX_MICROPHONE                                   = 3;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL                = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION                 = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY                 = 2;
        private const int INDEX_OUTPUT_SENSITIVITY_W_AUDIO_OUTPUT_SWITCH_LOW = 3;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT               = 6;

        Model1000X20AudioCable(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_MP3_PLAYER, data[COL_MP3_PLAYER_SN]));
            serialNo.Add(new Field(INDEX_AMP_PACK, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_3, data[COL_DRIVER3_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_4, data[COL_DRIVER4_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_5, data[COL_DRIVER5_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_6, data[COL_DRIVER6_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_7, data[COL_DRIVER7_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_OUTPUT_SENSITIVITY_W_AUDIO_OUTPUT_SWITCH_LOW, data[COL_AUDIO_OUTPUT_SW_OFF]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_SYSTEM_BACKGROUND_NOISE_OUTPUT]));
        }
    }

    class Model1000X30AudioCable : Model
    {
        private const int COL_MP3_PLAYER_SN                  = 7;
        private const int COL_ELEC_SN                        = 8;
        private const int COL_MIC_SN                         = 9;

        private const int COL_DRIVER1_SN                     = 10;
        private const int COL_DRIVER2_SN                     = 11;
        private const int COL_DRIVER3_SN                     = 12;
        private const int COL_DRIVER4_SN                     = 13;
        private const int COL_DRIVER5_SN                     = 14;
        private const int COL_DRIVER6_SN                     = 15;
        private const int COL_DRIVER7_SN                     = 16;

        private const int COL_MP3_SENSITIVITY                = 17;
        private const int COL_MAX_SPL                        = 18;
        private const int COL_VOL_FUNCTION                   = 19;
        private const int COL_MIC_SENSITIVITY                = 20;
        private const int COL_AUDIO_OUTPUT_SW_OFF            = 21;
        private const int COL_SYSTEM_BACKGROUND_NOISE_OUTPUT = 22;

        private const int COL_REMARKS                        = 26;

        private const int INDEX_MP3_PLAYER                                   = 1;
        private const int INDEX_AMP_PACK                                     = 2;
        private const int INDEX_MICROPHONE                                   = 3;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL                = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION                 = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY                 = 2;
        private const int INDEX_OUTPUT_SENSITIVITY_W_AUDIO_OUTPUT_SWITCH_LOW = 3;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT               = 6;

        Model1000X30AudioCable(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_MP3_PLAYER, data[COL_MP3_PLAYER_SN]));
            serialNo.Add(new Field(INDEX_AMP_PACK, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_3, data[COL_DRIVER3_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_4, data[COL_DRIVER4_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_5, data[COL_DRIVER5_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_6, data[COL_DRIVER6_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_7, data[COL_DRIVER7_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_OUTPUT_SENSITIVITY_W_AUDIO_OUTPUT_SWITCH_LOW, data[COL_AUDIO_OUTPUT_SW_OFF]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_SYSTEM_BACKGROUND_NOISE_OUTPUT]));
        }
    }

    class Model1000X100AudioCable : Model
    {
        private const int COL_MP3_PLAYER_SN                  = 7;
        private const int COL_ELEC_SN                        = 8;
        private const int COL_MIC_SN                         = 9;

        private const int COL_DRIVER1_SN                     = 10;
        private const int COL_DRIVER2_SN                     = 11;
        private const int COL_DRIVER3_SN                     = 12;
        private const int COL_DRIVER4_SN                     = 13;
        private const int COL_DRIVER5_SN                     = 14;
        private const int COL_DRIVER6_SN                     = 15;
        private const int COL_DRIVER7_SN                     = 16;

        private const int COL_MP3_SENSITIVITY                = 17;
        private const int COL_MAX_SPL                        = 18;
        private const int COL_VOL_FUNCTION                   = 19;
        private const int COL_MIC_SENSITIVITY                = 20;
        private const int COL_AUDIO_OUTPUT_SW_OFF            = 21;
        private const int COL_SYSTEM_BACKGROUND_NOISE_OUTPUT = 22;

        private const int COL_REMARKS                        = 26;

        private const int INDEX_MP3_PLAYER                                   = 1;
        private const int INDEX_AMP_PACK                                     = 2;
        private const int INDEX_MICROPHONE                                   = 3;

        private const int INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL                = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION                 = 1;
        private const int INDEX_MICROPHONE_INPUT_SENSITIVITY                 = 2;
        private const int INDEX_OUTPUT_SENSITIVITY_W_AUDIO_OUTPUT_SWITCH_LOW = 3;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT               = 6;

        Model1000X100AudioCable(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_MP3_PLAYER, data[COL_MP3_PLAYER_SN]));
            serialNo.Add(new Field(INDEX_AMP_PACK, data[COL_ELEC_SN]));
            serialNo.Add(new Field(INDEX_MICROPHONE, data[COL_MIC_SN]));

            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_3, data[COL_DRIVER3_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_4, data[COL_DRIVER4_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_5, data[COL_DRIVER5_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_6, data[COL_DRIVER6_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_7, data[COL_DRIVER7_SN]));

            subTests.Add(new Field(INDEX_MP3_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOL_FUNCTION]));
            subTests.Add(new Field(INDEX_MICROPHONE_INPUT_SENSITIVITY, data[COL_MIC_SENSITIVITY]));
            subTests.Add(new Field(INDEX_OUTPUT_SENSITIVITY_W_AUDIO_OUTPUT_SWITCH_LOW, data[COL_AUDIO_OUTPUT_SW_OFF]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTPUT, data[COL_SYSTEM_BACKGROUND_NOISE_OUTPUT]));
        }
    }

    class Model1000X2U : Model
    {
        private const int COL_2U_CHASSIS      = 7;

        private const int COL_DRIVER1_SN      = 8;
        private const int COL_DRIVER2_SN      = 9;
        private const int COL_DRIVER3_SN      = 10;
        private const int COL_DRIVER4_SN      = 11;
        private const int COL_DRIVER5_SN      = 12;
        private const int COL_DRIVER6_SN      = 13;
        private const int COL_DRIVER7_SN      = 14;

        private const int COL_MP3_SENSITIVITY = 15;
        private const int COL_MAX_SPL         = 16;
        private const int COL_VOLUME_FUNCTION = 17;
        private const int COL_MUTE_FUNCTION   = 18;
        private const int COL_OUTPUT_NOISE    = 19;

        private const int COL_REMARKS         = 23;

        private const int INDEX_2U_CHASSIS                     = 0;

        private const int INDEX_INPUT_SENSITIVITY_MAX_SPL      = 0;
        private const int INDEX_VOLUME_CONTROL_KNOB_FUNCTION   = 1;
        private const int INDEX_MUTE_FUNCTION                  = 2;
        private const int INDEX_SYSTEM_BACKGROUND_NOISE_OUTOUT = 6;

        Model1000X2U(string[] data) : base(data)
        {
            general.Add(new Field(INDEX_REMARKS, data[COL_REMARKS]));

            serialNo.Add(new Field(INDEX_2U_CHASSIS, data[COL_2U_CHASSIS]));
            
            driverSns.Add(new Field(INDEX_DRIVER_SN_1, data[COL_DRIVER1_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_2, data[COL_DRIVER2_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_3, data[COL_DRIVER3_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_4, data[COL_DRIVER4_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_5, data[COL_DRIVER5_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_6, data[COL_DRIVER6_SN]));
            driverSns.Add(new Field(INDEX_DRIVER_SN_7, data[COL_DRIVER7_SN]));

            subTests.Add(new Field(INDEX_INPUT_SENSITIVITY_MAX_SPL, Field.TwoPFEval(data[COL_MP3_SENSITIVITY], data[COL_MAX_SPL])));
            subTests.Add(new Field(INDEX_VOLUME_CONTROL_KNOB_FUNCTION, data[COL_VOLUME_FUNCTION]));
            subTests.Add(new Field(INDEX_MUTE_FUNCTION, data[COL_MUTE_FUNCTION]));
            subTests.Add(new Field(INDEX_SYSTEM_BACKGROUND_NOISE_OUTOUT, data[COL_OUTPUT_NOISE]));
        }
    }

    class Model1000Xi : Model
    {
        private const int COL_CONTROL_UNIT_SN = 7;
        private const int COL_ELEC_SN = 8;
        private const int COL_MIC_SN = 9;
        private const int COL_CARBON_FIBER_HD_SN = 10;
       
        private const int COL_DRIVER1_SN = 11;
        private const int COL_DRIVER2_SN = 12;
        private const int COL_DRIVER3_SN = 13;
        private const int COL_DRIVER4_SN = 14;
        private const int COL_DRIVER5_SN = 15;
        private const int COL_DRIVER6_SN = 16;
        private const int COL_DRIVER7_SN = 17;



        Model1000Xi(string[] data) : base(data)
        {

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
