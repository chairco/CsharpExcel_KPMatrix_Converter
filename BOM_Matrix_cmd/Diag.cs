#define DKey

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;

namespace BOM_Matrix_cmd
{
    class Diags
    {
        static string sProgramDescription = "Windows C# BOM_Matrix Utility";
        static string sBuildDate = "2014-08-28";
//        static string sRevision = "1.00a";

        static public void Logo(Boolean blNoLogo = false, Boolean blClearScreen = true)
        {
            string sRevision = "1.00d";
#if DKey            
            sRevision += "P";
#endif

            if (blNoLogo)
                return;

            if (blClearScreen)
                Command.clrscr();

            Command.PrintStr(Command.nSCREEN_LEFT(), 1, sProgramDescription, Command.nLOC_MIDDLE(), Command.bDEFAULT_COLOR());
            Command.PrintStr(Command.nSCREEN_LEFT(), 2, "Copyright by Pegatron, Build Date:" + sBuildDate + " Rev" + sRevision, Command.nLOC_MIDDLE(), Command.bDEFAULT_COLOR());
            Command.PrintStr(Command.nSCREEN_LEFT(), 2, "BG3_TPE ", Command.nLOC_RIGHT(), 0X08);
            Command.PrintCh(Command.nSCREEN_LEFT(), 3, '=', Command.bDEFAULT_COLOR(), Command.nSCREEN_RIGHT());

            return;
        }

        //-------------------------------------------------------------------------------------------------
        static public void ReadMe(Boolean blNoLogo = false)
        {
            Logo();

            int y = Command.wherey();

            if (!blNoLogo)
                y = 3;

            for (int i = 1; i <= 18; i++)
                Command.PrintCh(40, y + i, '|', 0x02, 1); // middle line
            Command.PrintCh(Command.nSCREEN_LEFT(), y + 19, '-', 0x08, Command.nSCREEN_RIGHT()); // bottom line
            Command.PrintCh(40, y + 19, '+', 0x08, 1); // intersection of bottom line and middle line

            Command.gotoxy(Command.nSCREEN_LEFT(), y + 1);
            // Basic function	
            int nY = 4;
            Command.PrintStr(0, nY++, "/?: readme", Command.nSCREEN_LEFT(), 0x07);
            Command.PrintStr(0, nY++, "Basic----------------------------------", Command.nSCREEN_LEFT(), 0x02);
            Command.PrintStr(0, nY++, "-nl: no display the logo and no clear", Command.nSCREEN_LEFT(), 0x07);
            Command.PrintStr(0, nY++, "     the screen", Command.nSCREEN_LEFT(), 0x07);
            //Command.PrintStr(0, nY++, "-erv <ErrorLevel>: Return Error Code", Command.nSCREEN_LEFT(), 0x07);
            Command.PrintStr(0, nY++, "BOM_Matrix function--------------------", Command.nSCREEN_LEFT(), 0x02);
            // Add other function descriptions here
            Command.PrintStr(0, nY++, "/C [FATP File] [MLB File]: Transfer", Command.nSCREEN_LEFT(), 0x07);
            Command.PrintStr(0, nY++, "BOM By Config.", Command.nSCREEN_LEFT(), 0x07);
            //Command.PrintStr(0, nY++, "/DELTA: Set the UAC to \"Never notify\".", Command.nSCREEN_LEFT(), 0x07);
            //Command.PrintStr(0, nY++, "/ChkClose: Check whether UAC is colsed.", Command.nSCREEN_LEFT(), 0x07);
            //Command.PrintStr(0, nY++, "/Default: Set the UAC to \"Default\".", Command.nSCREEN_LEFT(), 0x07);
            /*
            //RightWindow
            nY = 4;
            Command.PrintStr(41, nY++, "/T <dev>: Test Audio", Command.nSCREEN_RIGHT() - 40, 0x07);
            Command.gotoxy(Command.nSCREEN_LEFT(), y + 20);*/
            End(255);
            return;
        }

        //-------------------------------------------------------------------------------------------------
        static public void End(int nReturnCode, Boolean blNoLogo = false)
        {
            if (!blNoLogo)
            {
                int y = Command.wherey();
                if (y < 23)
                    y = 23;

                string sRetCode = "Return Code = ";
                sRetCode += nReturnCode.ToString();
                Command.PrintStr(0, y, sRetCode, Command.nLOC_MIDDLE(), 0x08);
            }
            Command.Cursor(true);
            Environment.Exit(nReturnCode);
            return;
        }

        static public bool blEncryption()
        {
#if DKey
            byte[] bData = null;
            if (!File.Exists("DK-diags.exe"))
                return false;
            using (FileStream fr = new FileStream("DK-diags.exe", FileMode.Open))
            {

                using (BinaryReader br = new BinaryReader(fr))
                {
                    if (fr.Length != 0x36E00)
                        return false;
                    bData = br.ReadBytes((int)fr.Length);
                    if (bData[0x221] != 0x72)
                    {
                        //                        Console.WriteLine("1:".ToString()+bData[0x221].ToString());
                        return false;
                    }
                    if (bData[0xd48] != 0x0A)
                    {
                        //                        Console.WriteLine("2:".ToString() + bData[0xd48].ToString());
                        return false;
                    }
                    if (bData[0x1000E] != 0x07)
                    {
                        //                        Console.WriteLine("3:".ToString() + bData[0x1000E].ToString());
                        return false;
                    }
                    if (bData[0x19BD4] != 0x16)
                    {
                        //                        Console.WriteLine("4:".ToString() + bData[0x19BD4].ToString());
                        return false;
                    }

                    if (bData[0x1DACC] != 0x01)
                    {
                        //                        Console.WriteLine("5:".ToString() + bData[0x1DACC].ToString());
                        return false;
                    }

                    if (bData[0x241E6] != 0x66)
                    {
                        //                        Console.WriteLine("6:".ToString() + bData[0x241E6].ToString());
                        return false;
                    }

                    if (bData[0x2F378] != 0x20)
                    {
                        //                        Console.WriteLine("7:".ToString() + bData[0x2F378].ToString());
                        return false;
                    }

                    if (bData[0x31AB8] != 0x2C)
                    {
                        //                        Console.WriteLine("8:".ToString() + bData[0x31AB8].ToString());
                        return false;
                    }
                }
            }
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.FileName = "DK-diags.exe";
            p.StartInfo.Arguments = "/t";
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.Start();
            p.WaitForExit();
            int returnCode = p.ExitCode;
            if (returnCode != 0)
                return false;
#endif
            return true;
        }
    }

    class Command
    {
        private const int STD_OUTPUT_HANDLE = -11;
        private const byte EMPTY = 32;

        private const int CREEN_TOP = 1;
        static public int nCREEN_TOP() { return CREEN_TOP; }
        private const int SCREEN_LEFT = 1;
        static public int nSCREEN_LEFT() { return SCREEN_LEFT; }
        private const int SCREEN_BOTTOM = 25;
        public int nSCREEN_BOTTOM() { return SCREEN_BOTTOM; }
        private const int SCREEN_RIGHT = 80;
        static public int nSCREEN_RIGHT() { return SCREEN_RIGHT; }
        private const int LOC_X = 0;
        static public int nLOC_X() { return LOC_X; }
        private const int LOC_LEFT = 1;
        static public int nLOC_LEFT() { return LOC_LEFT; }
        private const int LOC_MIDDLE = 2;
        static public int nLOC_MIDDLE() { return LOC_MIDDLE; }
        private const int LOC_RIGHT = 3;
        static public int nLOC_RIGHT() { return LOC_RIGHT; }
        private const byte DEFAULT_COLOR = 7;
        static public byte bDEFAULT_COLOR() { return DEFAULT_COLOR; }

        [StructLayout(LayoutKind.Sequential)]
        struct COORD
        {
            public short X;
            public short Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct SMALL_RECT
        {
            public short Left;
            public short Top;
            public short Right;
            public short Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct CONSOLE_SCREEN_BUFFER_INFO
        {
            public COORD dwSize;
            public COORD dwCursorPosition;
            public int wAttributes;
            public SMALL_RECT srWindow;
            public COORD dwMaximumWindowSize;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct CONSOLE_CURSOR_INFO
        {
            public uint dwSize;
            public int bVisible;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct CHAR_INFO
        {
            public char AsciiChar;
            public short Attributes;
        }

        [DllImport("kernel32.dll", EntryPoint = "GetStdHandle", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern int GetStdHandle(int nStdHandle);

        [DllImport("kernel32.dll", EntryPoint = "FillConsoleOutputCharacter", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern int FillConsoleOutputCharacter(int hConsoleOutput, short cCharacter, int nLength, COORD dwWriteCoord, ref uint lpNumberOfCharsWritten);

        [DllImport("kernel32.dll", EntryPoint = "FillConsoleOutputAttribute", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern int FillConsoleOutputAttribute(int hConsoleOutput, ushort cAttribute, uint nLength, COORD dwWriteCoord, ref uint lpNumberOfAttrsWritten);

        [DllImport("kernel32.dll", EntryPoint = "GetConsoleScreenBufferInfo", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern int GetConsoleScreenBufferInfo(int hConsoleOutput, ref CONSOLE_SCREEN_BUFFER_INFO lpConsoleScreenBufferInfo);

        [DllImport("kernel32.dll", EntryPoint = "SetConsoleCursorPosition", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern int SetConsoleCursorPosition(int hConsoleOutput, COORD dwCursorPosition);

        [DllImport("kernel32.dll", EntryPoint = "SetConsoleCursorInfo", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern int SetConsoleCursorInfo(int hConsoleOutput, ref CONSOLE_CURSOR_INFO lpConsoleCursorInfo);

        [DllImport("kernel32.dll", EntryPoint = "SetConsoleTextAttribute", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern int SetConsoleTextAttribute(int hConsoleOutput, ushort wAttributes);

        //[DllImport("kernel32.dll", EntryPoint = "WriteConsoleOutput", SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        //private static extern int WriteConsoleOutput(int hConsoleOutput, ref CHAR_INFO[] lpBuffer, COORD dwBufferSize, COORD dwBufferCoord, ref SMALL_RECT lpWriteRegion);
        [DllImport("kernel32.dll", EntryPoint = "WriteConsoleOutput", SetLastError = true, CharSet = CharSet.Unicode)]
        extern static bool WriteConsoleOutput(int handle, CHAR_INFO[] buffer, COORD bsize, COORD bpos, ref SMALL_RECT region);


        [DllImport("coredll.dll", EntryPoint = "DeviceIoControl", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool DeviceIoControl(IntPtr hDevice, uint dwIoControlCode, byte[] lpInBuffer, // LPVOID lpInBuffer - any input data requiredfor the IOCTL
           int nInBufferSize, byte[] lpOutBuffer, int nOutBufferSize, ref ulong lpBytesReturned,//ref int lpBytesReturned,
           IntPtr lpOverlapped);

        static int __BACKGROUND = 0;
        static int __FOREGROUND = 7;

        static COORD GetConsoleBuffer()
        {
            COORD coord;
            CONSOLE_SCREEN_BUFFER_INFO info = new CONSOLE_SCREEN_BUFFER_INFO();

            coord.X = 0;
            coord.Y = 0;
            if (GetConsoleScreenBufferInfo(
               GetStdHandle(STD_OUTPUT_HANDLE),
               ref info
               ) > 0)
            {
                coord = info.dwSize;
            }

            return coord;
        }

        //cursor goto xy 
        static public void gotoxy(int x, int y)
        {
            COORD c;

            c.X = (short)(x - 1);
            c.Y = (short)(y - 1);
            SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), c);
        }

        //clear screen 
        static public void clrscr()
        {
            UInt32 written = new UInt32();
            COORD coord;
            COORD buffer;
            uint nLength = 8000;
            buffer = GetConsoleBuffer();

            if (buffer.X > 0 || buffer.Y > 0)
            {
                nLength = (uint)(buffer.X * buffer.Y);
            }
            coord.X = 0;
            coord.Y = 0;
            FillConsoleOutputAttribute(GetStdHandle(STD_OUTPUT_HANDLE), (ushort)(__FOREGROUND + (__BACKGROUND << 4)), nLength, coord, ref written);
            FillConsoleOutputCharacter(GetStdHandle(STD_OUTPUT_HANDLE), (byte)' ', (int)nLength, coord, ref written);
            gotoxy(1, 1);
        }

        public void clreol()
        {
            COORD coord;
            UInt32 written = new UInt32();
            CONSOLE_SCREEN_BUFFER_INFO info = new CONSOLE_SCREEN_BUFFER_INFO();

            GetConsoleScreenBufferInfo(GetStdHandle(STD_OUTPUT_HANDLE), ref info);
            coord.X = info.dwCursorPosition.X;
            coord.Y = info.dwCursorPosition.Y;

            FillConsoleOutputCharacter(GetStdHandle(STD_OUTPUT_HANDLE), (byte)' ', info.dwSize.X - info.dwCursorPosition.X, coord, ref written);
            gotoxy(coord.X, coord.Y);
        }

        public void delline()
        {
            COORD coord;
            UInt32 written = new UInt32();
            CONSOLE_SCREEN_BUFFER_INFO info = new CONSOLE_SCREEN_BUFFER_INFO();

            GetConsoleScreenBufferInfo(GetStdHandle(STD_OUTPUT_HANDLE), ref info);
            coord.X = info.dwCursorPosition.X;
            coord.Y = info.dwCursorPosition.Y;

            FillConsoleOutputCharacter(GetStdHandle(STD_OUTPUT_HANDLE), (byte)' ', info.dwSize.X * info.dwCursorPosition.Y, coord, ref written);
            gotoxy(info.dwCursorPosition.X + 1, info.dwCursorPosition.Y + 1);
        }

        public void _setcursortype(int type)
        {
            CONSOLE_CURSOR_INFO Info = new CONSOLE_CURSOR_INFO();

            Info.dwSize = (uint)type;
            SetConsoleCursorInfo(GetStdHandle(STD_OUTPUT_HANDLE), ref Info);
        }

        public void Cursor(int type)
        {
            CONSOLE_CURSOR_INFO Info = new CONSOLE_CURSOR_INFO();

            Info.bVisible = type;
            SetConsoleCursorInfo(GetStdHandle(STD_OUTPUT_HANDLE), ref Info);
        }

        public void textattr(int _attr)
        {
            SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), (ushort)_attr);
        }

        public void textbackground(int color)
        {
            __BACKGROUND = color;
            SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), (ushort)(__FOREGROUND + (color << 4)));
        }

        public void textcolor(int color)
        {
            __FOREGROUND = color;
            SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), (ushort)(color + (__BACKGROUND << 4)));
        }

        static public int wherex()
        {
            CONSOLE_SCREEN_BUFFER_INFO info = new CONSOLE_SCREEN_BUFFER_INFO();

            GetConsoleScreenBufferInfo(GetStdHandle(STD_OUTPUT_HANDLE), ref info);
            return info.dwCursorPosition.X + 1;
        }

        static public int wherey()
        {
            CONSOLE_SCREEN_BUFFER_INFO info = new CONSOLE_SCREEN_BUFFER_INFO();

            GetConsoleScreenBufferInfo(GetStdHandle(STD_OUTPUT_HANDLE), ref info);
            return info.dwCursorPosition.Y + 1;
        }

        static public void PrintStr(int x, int y, string s, int location, byte color)
        {
            int print_x = x;
            int i = 0, j = 0;
            SMALL_RECT r = new SMALL_RECT();
            CHAR_INFO[] buffer = new CHAR_INFO[25 * 80];
            COORD c1, c2;

            switch (location)
            {
                case LOC_LEFT:
                    print_x = 1;
                    break;
                case LOC_MIDDLE:
                    print_x = (int)((82 - s.Length) / 2);
                    break;
                case LOC_RIGHT:
                    print_x = 82 - (int)s.Length;
                    break;
                case LOC_X:
                default:
                    break;
            }
            gotoxy(print_x + (int)s.Length, y);

            c1.X = 80;
            c1.Y = 25;
            c2.X = 0;
            c2.Y = 0;

            r.Left = (short)(print_x - 1);
            r.Bottom = (short)(y - 1);
            r.Right = (short)(r.Left + (int)s.Length - 1);
            r.Top = (short)(y - 1);

            for (j = 0; j < s.Length; j++)
            {
                buffer[i * 25 + j].AsciiChar = s[j];
                buffer[i * 25 + j].Attributes = color;
            }

            WriteConsoleOutput(GetStdHandle(STD_OUTPUT_HANDLE), buffer, c1, c2, ref r);
        }

        static public void PrintCh(int x, int y, char c, byte color, int count)
        {
            int i = 0, j = 0, n = 0;
            SMALL_RECT r = new SMALL_RECT();
            CHAR_INFO[] buffer = new CHAR_INFO[25 * 80];
            COORD c1, c2;

            c1.X = 80;
            c1.Y = 25;
            c2.X = 0;
            c2.Y = 0;

            r.Left = (short)(x - 1);
            r.Bottom = (short)(y - 1);
            r.Right = (short)(r.Left + count - 1);
            r.Top = (short)(y - 1);
            gotoxy((x + count) % 80, y + (x + count) / 80);

            for (j = 0; j < count; j++)
            {
                buffer[i * 25 + j].AsciiChar = c;
                buffer[i * 25 + j].Attributes = color;
                n++;
            }

            WriteConsoleOutput(GetStdHandle(STD_OUTPUT_HANDLE), buffer, c1, c2, ref r);
        }

        static public void Cursor(Boolean type)
        {
            CONSOLE_CURSOR_INFO Info = new CONSOLE_CURSOR_INFO();

            Info.bVisible = type ? 1 : 0;
            SetConsoleCursorInfo(GetStdHandle(STD_OUTPUT_HANDLE), ref Info);
        }

        public UInt32 atoh(string s)
        {
            UInt32 num = UInt32.Parse(s, System.Globalization.NumberStyles.HexNumber);
            return num;
        }

        public Boolean blParameter(string s)
        {
            if (s.Length == 0) return false;
            if (s[0] == '-' || s[0] == '/' || s[0] == 0x20 || s[0] == 0x00)
                return false;
            return true;
        }
        public Boolean blDec(string s)
        {
            int n = 0;
            for (n = 0; n < s.Length; n++)
            {
                if (!Char.IsDigit(s, n))
                    break;
            }
            if (n > 0 && n == s.Length)
                return true;
            return false;
        }
    }
}
