using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;   // For the DllImport of AIOWDM Driver API
using System.Text;

namespace AIOWDMNet
{
    public class AIOWDM
    {
        // Prototype AIOWDM.dll data and functions in managed C# for Dot Net
        // Using C# data types and default marshaling of unmanaged code

        // Predefined constants for device index:
        public const UInt32 diNone = 0xFFFFFFFF;    // "-1"
        public const UInt32 diFirst = 0xFFFFFFFE;   // "-2"  First board
        public const UInt32 diOnly = 0xFFFFFFFD;    // "-3"  One and only board 

        public const UInt32 TIME_METHOD_NOW = 0;
        public const UInt32 TIME_METHOD_WAIT_INPUT_ENABLE = 86;
        public const UInt32 TIME_METHOD_NOW_AND_ABORT = 5;
        public const UInt32 TIME_METHOD_WHEN_INPUT_ENABLE = 1;

        // Watchdog function constants:
        public const UInt32 WDG_ACTION_IGNORE = 0;
        public const UInt32 WDG_ACTION_DISABLE = 1;
        public const UInt32 WDG_ACTION_SOFT_RESTART = 2;
        public const UInt32 WDG_ACTION_MOSTLY_SOFT_RESTART = 4;
        public const Double PCI_WDG_CSM_RATE = 2.08333;
        public const Double P104_WDG_CSM_RATE = 2.08333;
        public const Double ISA_WDG_CSM_RATE = 0.894886;

        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern Int32 GetNumCards();

        // QueryDeviceInfo() Array of char Name version :
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 QueryCardInfo(Int32 CardNum, out UInt32 pDeviceID, out UInt32 pBase, out UInt32 pNameSize, [In, Out] Char[] Name);
        // QueryDeviceInfo() StringBuilder Name version :
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 QueryCardInfo(Int32 CardNum, out UInt32 pDeviceID, out UInt32 pBase, out UInt32 pNameSize, StringBuilder Name);
        // QueryDeviceInfo() String Name version :
        public static UInt32 QueryCardInfo(Int32 CardNum, out UInt32 pDeviceID, out UInt32 pBase, out UInt32 pNameSize, out String Name)
        {
            UInt32 Status;
            UInt32 NameSize = 256; // must pass this in by reference if you want name to be modified

            Char[] charName = new Char[NameSize];
            Status = QueryCardInfo(CardNum, out pDeviceID, out pBase, out NameSize, charName);

            Name = new String(charName, 0, (Int32)NameSize);
            pNameSize = NameSize;
            return Status;
        }


        // We need SetLastError = true for the two wait functions so they are "allowed" to set it.
        // Its not neccesarily an error it is used like a thread return value for already pending wait or intentional abort
        // This actually tels the CLR to call GetLastError immediatly after a call and save the value dont overwrite it.
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 QueryBARBase(Int32 CardNum, UInt32 BARIndex, out UInt32 pBase);

        [DllImport("AIOWDM.DLL", SetLastError = true, CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 WaitForIRQ(Int32 CardNum);

        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 AbortRequest(Int32 CardNum);

        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 CloseCard(Int32 CardNum);
        //[DllImport("AIOWDM.DLL", SetLastError = true, CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        //public static extern UInt32 COSWaitForIRQ(Int64 CardNum, UInt32 PPIs, void* pData); // ref Void ???
        // uses a void pointer to an array of 3 unsigned char or UInt16

        [DllImport("AIOWDM.DLL", SetLastError = true, CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 COSWaitForIRQ(Int32 CardNum, UInt32 PPIs, [In, Out] UInt16[] pData); // ref Void ???

        // Byte version:  COS boards will only have 1 or 2 PPIs so 3 or 6 bytes for chan ABC
        [DllImport("AIOWDM.DLL", SetLastError = true, CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 COSWaitForIRQ(Int32 CardNum, UInt32 PPIs, [In, Out] Byte[] pData); // ref Void ???

        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 WDGInit(Int32 CardNum);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 WDGHandleIRQ(Int32 CardNum, UInt32 Action);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern Double WDGSetTimeout(Int32 CardNum, Double Milliseconds, Double MHzClockRate);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern Double WDGSetResetDuration(Int32 CardNum, Double Milliseconds, Double MHzClockRate);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 WDGPet(Int32 CardNum);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern Double WDGReadTemp(Int32 CardNum);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 WDGReadStatus(Int32 CardNum);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 WDGStart(Int32 CardNum);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 WDGStop(Int32 CardNum);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 EmergencyReboot();
        //UInt32 EmergencyReboot(void);

        // We use an 8 bit byte here but return/pass a 16 bit byte to accomidate possible AA55 error return for value:
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 InPortB(UInt32 Port);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 OutPortB(UInt32 Port, Byte Value);

        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 InPort(UInt32 Port);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 OutPort(UInt32 Port, UInt16 Value);

        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 InPortL(UInt32 Port);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 OutPortL(UInt32 Port, UInt32 Value);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 InPortDWord(UInt32 Port);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 OutPortDWord(UInt32 Port, UInt32 Value);

        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 RelInPortB(Int32 CardNum, UInt32 Port);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 RelOutPortB(Int32 CardNum, UInt32 Port, Byte Value);

        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 RelInPort(Int32 CardNum, UInt32 Port);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 RelOutPort(Int32 CardNum, UInt32 Port, UInt16 Value);

        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 RelInPortL(Int32 CardNum, UInt32 Port);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 RelOutPortL(Int32 CardNum, UInt32 Port, UInt32 Value);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 RelInPortDWord(Int32 CardNum, UInt32 Port);
        [DllImport("AIOWDM.DLL", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt16 RelOutPortDWord(Int32 CardNum, UInt32 Port, UInt32 Value);


    }
}

namespace Final_Audio_Test
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
