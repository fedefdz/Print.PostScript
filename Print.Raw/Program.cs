using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Print.Raw
{    
    public class RawPrinterAnsi
    {
        #region Structure and API declarions ...
        
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        private class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
        }

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StartDocPrinter(IntPtr hPrinter, int level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true,  CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

        [DllImport("winspool.drv", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int size);

        #endregion

        #region Private static Methods ...

        // SendBytesToPrinter()
        // When the function is given a printer name and an unmanaged array
        // of bytes, the function sends those bytes to the print queue.
        // Returns true on success, false on failure.
        private static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, int dwCount)
        {
            int dwError = 0, dwWritten = 0;
            IntPtr hPrinter = new IntPtr(0);
            DOCINFOA di = new DOCINFOA();
            bool bSuccess = false; // Assume failure unless you specifically succeed.
            di.pDocName = "My C#.NET RAW Document";
            di.pDataType = "RAW";

            // Open the printer.
            if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
            {
                // Start a document.
                if (StartDocPrinter(hPrinter, 1, di))
                {
                    // Start a page.
                    if (StartPagePrinter(hPrinter))
                    {
                        // Write your bytes.
                        bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                        EndPagePrinter(hPrinter);
                    }
                    EndDocPrinter(hPrinter);
                }
                ClosePrinter(hPrinter);
            }
            // If you did not succeed, GetLastError may give more information
            // about why not.
            if (bSuccess == false)
            {
                dwError = Marshal.GetLastWin32Error();
            }
            return bSuccess;
        }

        private static bool SendFileToPrinter(string szPrinterName, string szFileName)
        {
            // Open the file.
            FileStream fs = new FileStream(szFileName, FileMode.Open);
            // Create a BinaryReader on the file.
            BinaryReader br = new BinaryReader(fs);
            // Dim an array of bytes big enough to hold the file's contents.
            byte[] bytes = new byte[fs.Length];
            bool bSuccess = false;
            // Your unmanaged pointer.
            IntPtr pUnmanagedBytes = new IntPtr(0);
            int nLength;

            nLength = Convert.ToInt32(fs.Length);
            // Read the contents of the file into the array.
            bytes = br.ReadBytes(nLength);
            // Allocate some unmanaged memory for those bytes.
            pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
            // Copy the managed byte array into the unmanaged array.
            Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);
            // Send the unmanaged bytes to the printer.
            bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, nLength);
            // Free the unmanaged memory that you allocated earlier.
            Marshal.FreeCoTaskMem(pUnmanagedBytes);
            return bSuccess;
        }

        private static bool SendStringToPrinter(string szPrinterName, string szString)
        {
            IntPtr pBytes;
            int dwCount;

            // How many characters are in the string?
            // Fix from Nicholas Piasecki:
            // dwCount = szString.Length;
            dwCount = (szString.Length + 1) * Marshal.SystemMaxDBCSCharSize;

            // Assume that the printer is expecting ANSI text, and then convert
            // the string to ANSI text.
            pBytes = Marshal.StringToCoTaskMemAnsi(szString);            

            // Send the converted ANSI string to the printer.
            SendBytesToPrinter(szPrinterName, pBytes, dwCount);
            Marshal.FreeCoTaskMem(pBytes);
            return true;
        }
        
        #endregion

        #region ESC/POS Commands ...

        // ESC/POS Command Set for Reliance Thermal Printer
        // http://reliance-escpos-commands.readthedocs.io/en/latest/index.html
        public static readonly string Initialize = Convert.ToString((char)27) + Convert.ToString((char)64);
        public static readonly string InitializeFull = Convert.ToString((char)27) + Convert.ToString((char)64)
            + Convert.ToString((char)27) + Convert.ToString((char)50)
            + Convert.ToString((char)27) + Convert.ToString((char)123) + Convert.ToString(0)
            + Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)19)
            + Convert.ToString((char)27) + Convert.ToString((char)82) + Convert.ToString((char)7);

        public static readonly string FontCodePage0 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)0);
        public static readonly string FontCodePage2 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)2);
        public static readonly string FontCodePage3 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)3);
        public static readonly string FontCodePage4 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)4);
        public static readonly string FontCodePage5 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)5);
        public static readonly string FontCodePage17 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)17);
        public static readonly string FontCodePage18 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)18);
        public static readonly string FontCodePage19 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)19);
        public static readonly string FontCodePage255 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)255);

        public static readonly string CutPaper = Convert.ToString((char)27) + Convert.ToString((char)105);
        public static readonly string FontDefault = Convert.ToString((char)27) + Convert.ToString((char)77) + Convert.ToString((char)0);
        public static readonly string FontSmall = Convert.ToString((char)27) + Convert.ToString((char)77) + Convert.ToString((char)1);
        public static readonly string AlignCenter = Convert.ToString((char)27) + Convert.ToString((char)97) + Convert.ToString((char)1);
        public static readonly string AlignRight = Convert.ToString((char)27) + Convert.ToString((char)97) + Convert.ToString((char)2);
        public static readonly string AlignLeft = Convert.ToString((char)27) + Convert.ToString((char)97) + Convert.ToString((char)0);
        public static readonly string Underline = Convert.ToString((char)27) + Convert.ToString((char)45) + Convert.ToString((char)1);
        public static readonly string NoUnderline = Convert.ToString((char)27) + Convert.ToString((char)45) + Convert.ToString((char)0);
        public static readonly string Bold = Convert.ToString((char)27) + Convert.ToString((char)69) + Convert.ToString((char)1);
        public static readonly string NoBold = Convert.ToString((char)27) + Convert.ToString((char)69) + Convert.ToString((char)0);
        public static readonly string SizeH1 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)0);
        public static readonly string SizeH2 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)17);
        public static readonly string SizeH4 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)51);
        public static readonly string SizeH8 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)119);
        public static readonly string NewLine = Convert.ToString((char)10);
        public static readonly string BarCodeHeightX20 = Convert.ToString((char)29) + Convert.ToString((char)104) + Convert.ToString((char)20);
        public static readonly string BarCodeBelow = Convert.ToString((char)29) + Convert.ToString((char)72) + Convert.ToString((char)2);
        public static readonly string BarCodeFormA = Convert.ToString((char)29) + Convert.ToString((char)107) + Convert.ToString((char)67) + Convert.ToString((char)13);
        public static readonly string BarCodeFormB = Convert.ToString((char)29) + Convert.ToString((char)107) + Convert.ToString((char)2);
        #endregion

        public string Text { get; private set; }
        public string PrinterName { get; set; }

        public RawPrinterAnsi()
        {
            Text = string.Empty;
            PrinterName = GetDefaultPrinter();
        }

        public string NormalizeCharactersAscii(string text)
        {
            //text = text.Replace('Á', (char) 181).Replace("&Aacute;", Convert.ToString((char) 181))
            //    .Replace('á', (char) 160).Replace("&aacute;", Convert.ToString((char) 160))
            //    .Replace('À', (char) 183).Replace("&Agrave;", Convert.ToString((char) 181))
            //    .Replace('à', (char) 133).Replace("&agrave;", Convert.ToString((char) 133))
            //    .Replace('Ä', (char) 142).Replace("&Auml;", Convert.ToString((char) 142))
            //    .Replace('ä', (char) 132).Replace("&auml;", Convert.ToString((char) 132))

            //    .Replace('É', (char) 144).Replace("&Eacute;", Convert.ToString((char) 144))
            //    .Replace('é', (char) 130).Replace("&eacute;", Convert.ToString((char) 130))
            //    .Replace('È', (char) 212).Replace("&Egrave;", Convert.ToString((char) 212))
            //    .Replace('è', (char) 138).Replace("&egrave;", Convert.ToString((char) 138))
            //    .Replace('Ë', (char) 211).Replace("&Euml;", Convert.ToString((char) 211))
            //    .Replace('ë', (char) 137).Replace("&euml;", Convert.ToString((char) 137))

            //    .Replace('Í', (char) 214).Replace("&Iacute;", Convert.ToString((char) 214))
            //    .Replace('í', (char) 161).Replace("&iacute;", Convert.ToString((char) 161))
            //    .Replace('Ì', (char) 222).Replace("&Igrave;", Convert.ToString((char) 222))
            //    .Replace('ì', (char) 141).Replace("&igrave;", Convert.ToString((char) 141))
            //    .Replace('Ï', (char) 216).Replace("&Iuml;", Convert.ToString((char) 216))
            //    .Replace('ï', (char) 139).Replace("&iuml;", Convert.ToString((char) 139))

            //    .Replace('Ó', (char) 224).Replace("&Oacute;", Convert.ToString((char) 224))
            //    .Replace('ó', (char) 162).Replace("&oacute;", Convert.ToString((char) 162))
            //    .Replace('Ò', (char) 227).Replace("&Ograve;", Convert.ToString((char) 227))
            //    .Replace('ò', (char) 149).Replace("&ograve;", Convert.ToString((char) 149))
            //    .Replace('Ö', (char) 153).Replace("&Ouml;", Convert.ToString((char) 153))
            //    .Replace('ö', (char) 148).Replace("&ouml;", Convert.ToString((char) 148))

            //    .Replace('Ú', (char)218).Replace("&Uacute;", Convert.ToString((char) 233))
            //    .Replace('ú', (char) 163).Replace("&uacute;", Convert.ToString((char) 163))
            //    .Replace('Ù', (char) 235).Replace("&Ugrave;", Convert.ToString((char) 235))
            //    .Replace('ù', (char) 151).Replace("&ugrave;", Convert.ToString((char) 151))
            //    .Replace('Ü', (char) 154).Replace("&Uuml;", Convert.ToString((char) 154))
            //    .Replace('ü', (char) 129).Replace("&uuml;", Convert.ToString((char) 129))

            //    .Replace('Ñ', (char) 165).Replace("&Ntilde;", Convert.ToString((char) 165))
            //    .Replace('ñ', (char) 164).Replace("&ntilde;", Convert.ToString((char) 164))

            //    .Replace('Ç', (char) 128).Replace("&Ccedil;", Convert.ToString((char) 128))
            //    .Replace('ç', (char) 135).Replace("&ccedil;", Convert.ToString((char) 135))

            //    .Replace('€', (char) 213).Replace("&euro;", Convert.ToString((char) 213))

            //    .Replace('º', (char) 167).Replace("&ordm;", Convert.ToString((char) 167))
            //    .Replace('ª', (char) 166).Replace("&ordf;", Convert.ToString((char) 166));
         
            byte[] bytes = Encoding.Default.GetBytes(text);
            text = Encoding.ASCII.GetString(bytes);
            return text;            
        }

        public string NormalizeCharactersUtf8(string text)
        {
            byte[] bytes = Encoding.Default.GetBytes(text);
            text = Encoding.UTF8.GetString(bytes);
            return text;
        }

        public string NormalizeCharactersUnicode(string text)
        {
            byte[] bytes = Encoding.Default.GetBytes(text);
            text = Encoding.UTF8.GetString(bytes);
            return text;
        }

        public string GetDefaultPrinter()
        {
            PrintDocument pd = new PrintDocument();
            StringBuilder dp = new StringBuilder(256);
            int size = dp.Capacity;
            if (GetDefaultPrinter(dp, ref size))
            {
                pd.PrinterSettings.PrinterName = dp.ToString().Trim();
            }
            return pd.PrinterSettings.PrinterName;
        }

        public void Clear()
        {
            Text = string.Empty;
        }

        public void Draw(string text)
        {
            Text += text;
        }

        public void Draw(string text, Align align, uint width)
        {
            string textDrawing = string.Empty;

            switch (align)
            {
                case Align.Center:
                    var margin = new string (' ', (int)Math.Truncate(width - text.Length / (decimal)2));
                    textDrawing = margin + text + margin;
                    break;
                case Align.Left:
                    var left = new string(' ', (int) width - text.Length);
                    textDrawing = text + left;
                    break;
                case Align.Right:
                    var right = new string(' ', (int)width - text.Length);
                    textDrawing = right + text;
                    break;                
            }
            
            Text += textDrawing;            
        }

        public void DrawLine(string text)
        {
            Text += NewLine + text;
        }

        public void DrawLine(string text, Align align)
        {
            switch (align)
            {
                case Align.Center:
                    Text += NewLine + AlignCenter + text;
                    break;
                case Align.Left:
                    Text += NewLine + AlignLeft + text;
                    break;
                case Align.Right:
                    Text += NewLine + AlignRight + text;
                    break;                
            }                        
        }

        public void DrawLine(string text, Align align, uint width)
        {
            Draw(NewLine);
            Draw(text, align, width);            
        }

        public void SetAlign(Align align)
        {
            switch (align)
            {
                case Align.Center:
                    Text += NewLine + AlignCenter;
                    break;
                case Align.Left:
                    Text += NewLine + AlignLeft;
                    break;
                case Align.Right:
                    Text += NewLine + AlignRight;
                    break;
            }
        }

        public void BarCodeModelA(string barcode)
        {
            Text += NewLine + BarCodeHeightX20 + BarCodeBelow + BarCodeFormA + barcode + NewLine;
        }

        public void BarCodeModelB(string barcode)
        {
            Text += NewLine + BarCodeHeightX20 + BarCodeBelow + BarCodeFormB + barcode + NewLine;
        }

        public void BarCode(string bardcode, int model)
        {
            if (model == 765 || model == 605)            
                BarCodeModelA(bardcode);
            else
                BarCodeModelB(bardcode);                            
        }

        public void Print()
        {
            string text = Initialize + FontCodePage0 + "RawPrinterAnsi" + NewLine + "CodePage 0" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage2 + "RawPrinterAnsi" + NewLine + "CodePage 2" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage3 + "RawPrinterAnsi" + NewLine + "CodePage 3" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage4 + "RawPrinterAnsi" + NewLine + "CodePage 4" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage5 + "RawPrinterAnsi" + NewLine + "CodePage 5" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage17 + "RawPrinterAnsi" + NewLine + "CodePage 17" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage18 + "RawPrinterAnsi" + NewLine + "CodePage 18" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage19 + "RawPrinterAnsi" + NewLine + "CodePage 19" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage255 + "RawPrinterAnsi" + NewLine + "CodePage 255" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = InitializeFull + "RawPrinterAnsi" + NewLine + "FULL" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            // ------------------------------------------------------------

            text = Initialize + FontCodePage0 + "RawPrinterAnsi" + NewLine + "CodePage 0" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage2 + "RawPrinterAnsi" + NewLine + "CodePage 2" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage3 + "RawPrinterAnsi" + NewLine + "CodePage 3" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage4 + "RawPrinterAnsi" + NewLine + "CodePage 4" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage5 + "RawPrinterAnsi" + NewLine + "CodePage 5" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage17 + "RawPrinterAnsi" + NewLine + "CodePage 17" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage18 + "RawPrinterAnsi" + NewLine + "CodePage 18" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage19 + "RawPrinterAnsi" + NewLine + "CodePage 19" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage255 + "RawPrinterAnsi" + NewLine + "CodePage 255" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = InitializeFull + "RawPrinterAnsi" + NewLine + "FULL" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            // -----------------------------------------------------------------

            text = Initialize + FontCodePage0 + "RawPrinterAnsi" + NewLine + "CodePage 0" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage2 + "RawPrinterAnsi" + NewLine + "CodePage 2" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage3 + "RawPrinterAnsi" + NewLine + "CodePage 3" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage4 + "RawPrinterAnsi" + NewLine + "CodePage 4" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage5 + "RawPrinterAnsi" + NewLine + "CodePage 5" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage17 + "RawPrinterAnsi" + NewLine + "CodePage 17" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage18 + "RawPrinterAnsi" + NewLine + "CodePage 18" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage19 + "RawPrinterAnsi" + NewLine + "CodePage 19" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage255 + "RawPrinterAnsi" + NewLine + "CodePage 255" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = InitializeFull + "RawPrinterAnsi" + NewLine + "FULL" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            // ------------------------------------------------

            text = Initialize + FontCodePage0 + "RawPrinterAnsi" + NewLine + "CodePage 0" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage2 + "RawPrinterAnsi" + NewLine + "CodePage 2" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage3 + "RawPrinterAnsi" + NewLine + "CodePage 3" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage4 + "RawPrinterAnsi" + NewLine + "CodePage 4" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage5 + "RawPrinterAnsi" + NewLine + "CodePage 5" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage17 + "RawPrinterAnsi" + NewLine + "CodePage 17" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage18 + "RawPrinterAnsi" + NewLine + "CodePage 18" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage19 + "RawPrinterAnsi" + NewLine + "CodePage 19" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage255 + "RawPrinterAnsi" + NewLine + "CodePage 255" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = InitializeFull + "RawPrinterAnsi" + NewLine + "FULL" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);
        }

        public enum Align
        {
            Center,
            Left,
            Right
        }
    }

    public class RawPrinterUnicode
    {
        #region Structure and API declarions ...

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
        }

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StartDocPrinter(IntPtr hPrinter, int level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

        [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int size);

        #endregion

        #region Private static Methods ...

        // SendBytesToPrinter()
        // When the function is given a printer name and an unmanaged array
        // of bytes, the function sends those bytes to the print queue.
        // Returns true on success, false on failure.
        private static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, int dwCount)
        {
            int dwError = 0, dwWritten = 0;
            IntPtr hPrinter = new IntPtr(0);
            DOCINFOA di = new DOCINFOA();
            bool bSuccess = false; // Assume failure unless you specifically succeed.
            di.pDocName = "My C#.NET RAW Document";
            di.pDataType = "RAW";

            // Open the printer.
            if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
            {
                // Start a document.
                if (StartDocPrinter(hPrinter, 1, di))
                {
                    // Start a page.
                    if (StartPagePrinter(hPrinter))
                    {
                        // Write your bytes.
                        bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                        EndPagePrinter(hPrinter);
                    }
                    EndDocPrinter(hPrinter);
                }
                ClosePrinter(hPrinter);
            }
            // If you did not succeed, GetLastError may give more information
            // about why not.
            if (bSuccess == false)
            {
                dwError = Marshal.GetLastWin32Error();
            }
            return bSuccess;
        }

        private static bool SendFileToPrinter(string szPrinterName, string szFileName)
        {
            // Open the file.
            FileStream fs = new FileStream(szFileName, FileMode.Open);
            // Create a BinaryReader on the file.
            BinaryReader br = new BinaryReader(fs);
            // Dim an array of bytes big enough to hold the file's contents.
            byte[] bytes = new byte[fs.Length];
            bool bSuccess = false;
            // Your unmanaged pointer.
            IntPtr pUnmanagedBytes = new IntPtr(0);
            int nLength;

            nLength = Convert.ToInt32(fs.Length);
            // Read the contents of the file into the array.
            bytes = br.ReadBytes(nLength);
            // Allocate some unmanaged memory for those bytes.
            pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
            // Copy the managed byte array into the unmanaged array.
            Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);
            // Send the unmanaged bytes to the printer.
            bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, nLength);
            // Free the unmanaged memory that you allocated earlier.
            Marshal.FreeCoTaskMem(pUnmanagedBytes);
            return bSuccess;
        }

        private static bool SendStringToPrinter(string szPrinterName, string szString)
        {
            IntPtr pBytes;
            int dwCount;

            // How many characters are in the string?
            // Fix from Nicholas Piasecki:
            // dwCount = szString.Length;
            dwCount = (szString.Length + 1) * Marshal.SystemMaxDBCSCharSize;
            
            // Assume that the printer is expecting ANSI text, and then convert
            // the string to Unicode text.
            pBytes = Marshal.StringToCoTaskMemUni(szString);

            // Send the converted ANSI string to the printer.
            SendBytesToPrinter(szPrinterName, pBytes, dwCount);
            Marshal.FreeCoTaskMem(pBytes);
            return true;
        }

        #endregion

        #region ESC/POS Commands ...

        // ESC/POS Command Set for Reliance Thermal Printer
        // http://reliance-escpos-commands.readthedocs.io/en/latest/index.html
        public static readonly string Initialize = Convert.ToString((char)27) + Convert.ToString((char)64);
        public static readonly string InitializeFull = Convert.ToString((char)27) + Convert.ToString((char)64) 
            + Convert.ToString((char)27) + Convert.ToString((char)50)
            + Convert.ToString((char)27) + Convert.ToString((char)123) + Convert.ToString(0) 
            + Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)19) 
            + Convert.ToString((char)27) + Convert.ToString((char)82) + Convert.ToString((char)7);

        public static readonly string FontCodePage0 = Convert.ToString((char) 27) + Convert.ToString((char) 116) + Convert.ToString((char)0);
        public static readonly string FontCodePage2 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)2);
        public static readonly string FontCodePage3 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)3);
        public static readonly string FontCodePage4 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)4);
        public static readonly string FontCodePage5 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)5);
        public static readonly string FontCodePage17 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)17);
        public static readonly string FontCodePage18 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)18);
        public static readonly string FontCodePage19 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)19);
        public static readonly string FontCodePage255 = Convert.ToString((char)27) + Convert.ToString((char)116) + Convert.ToString((char)255);

        public static readonly string CutPaper = Convert.ToString((char)27) + Convert.ToString((char)105);
        public static readonly string FontDefault = Convert.ToString((char)27) + Convert.ToString((char)77) + Convert.ToString((char)0);
        public static readonly string FontSmall = Convert.ToString((char)27) + Convert.ToString((char)77) + Convert.ToString((char)1);
        public static readonly string AlignCenter = Convert.ToString((char)27) + Convert.ToString((char)97) + Convert.ToString((char)1);
        public static readonly string AlignRight = Convert.ToString((char)27) + Convert.ToString((char)97) + Convert.ToString((char)2);
        public static readonly string AlignLeft = Convert.ToString((char)27) + Convert.ToString((char)97) + Convert.ToString((char)0);
        public static readonly string Underline = Convert.ToString((char)27) + Convert.ToString((char)45) + Convert.ToString((char)1);
        public static readonly string NoUnderline = Convert.ToString((char)27) + Convert.ToString((char)45) + Convert.ToString((char)0);
        public static readonly string Bold = Convert.ToString((char)27) + Convert.ToString((char)69) + Convert.ToString((char)1);
        public static readonly string NoBold = Convert.ToString((char)27) + Convert.ToString((char)69) + Convert.ToString((char)0);
        public static readonly string SizeH1 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)0);
        public static readonly string SizeH2 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)17);
        public static readonly string SizeH4 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)51);
        public static readonly string SizeH8 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)119);
        public static readonly string NewLine = Convert.ToString((char)10);
        public static readonly string BarCodeHeightX20 = Convert.ToString((char)29) + Convert.ToString((char)104) + Convert.ToString((char)20);
        public static readonly string BarCodeBelow = Convert.ToString((char)29) + Convert.ToString((char)72) + Convert.ToString((char)2);
        public static readonly string BarCodeFormA = Convert.ToString((char)29) + Convert.ToString((char)107) + Convert.ToString((char)67) + Convert.ToString((char)13);
        public static readonly string BarCodeFormB = Convert.ToString((char)29) + Convert.ToString((char)107) + Convert.ToString((char)2);
        #endregion

        public string Text { get; private set; }
        public string PrinterName { get; set; }

        public RawPrinterUnicode()
        {
            Text = string.Empty;
            PrinterName = GetDefaultPrinter();
        }

        public string NormalizeCharactersAscii(string text)
        {
            byte[] bytes = Encoding.Default.GetBytes(text);
            text = Encoding.ASCII.GetString(bytes);
            return text;            
        }

        public string NormalizeCharactersUtf8(string text)
        {
            byte[] bytes = Encoding.Default.GetBytes(text);
            text = Encoding.UTF8.GetString(bytes);
            return text;
        }

        public string NormalizeCharactersUnicode(string text)
        {
            byte[] bytes = Encoding.Default.GetBytes(text);
            text = Encoding.UTF8.GetString(bytes);
            return text;
        }

        public string GetDefaultPrinter()
        {
            PrintDocument pd = new PrintDocument();
            StringBuilder dp = new StringBuilder(256);
            int size = dp.Capacity;
            if (GetDefaultPrinter(dp, ref size))
            {
                pd.PrinterSettings.PrinterName = dp.ToString().Trim();
            }
            return pd.PrinterSettings.PrinterName;
        }

        public void Clear()
        {
            Text = string.Empty;
        }

        public void Draw(string text)
        {
            Text += text;
        }

        public void Draw(string text, Align align, uint width)
        {
            string textDrawing = string.Empty;

            switch (align)
            {
                case Align.Center:
                    var margin = new string(' ', (int)Math.Truncate(width - text.Length / (decimal)2));
                    textDrawing = margin + text + margin;
                    break;
                case Align.Left:
                    var left = new string(' ', (int)width - text.Length);
                    textDrawing = text + left;
                    break;
                case Align.Right:
                    var right = new string(' ', (int)width - text.Length);
                    textDrawing = right + text;
                    break;
            }

            Text += textDrawing;
        }

        public void DrawLine(string text)
        {
            Text += NewLine + text;
        }

        public void DrawLine(string text, Align align)
        {
            switch (align)
            {
                case Align.Center:
                    Text += NewLine + AlignCenter + text;
                    break;
                case Align.Left:
                    Text += NewLine + AlignLeft + text;
                    break;
                case Align.Right:
                    Text += NewLine + AlignRight + text;
                    break;
            }
        }

        public void DrawLine(string text, Align align, uint width)
        {
            Draw(NewLine);
            Draw(text, align, width);
        }

        public void SetAlign(Align align)
        {
            switch (align)
            {
                case Align.Center:
                    Text += NewLine + AlignCenter;
                    break;
                case Align.Left:
                    Text += NewLine + AlignLeft;
                    break;
                case Align.Right:
                    Text += NewLine + AlignRight;
                    break;
            }
        }

        public void BarCodeModelA(string barcode)
        {
            Text += NewLine + BarCodeHeightX20 + BarCodeBelow + BarCodeFormA + barcode + NewLine;
        }

        public void BarCodeModelB(string barcode)
        {
            Text += NewLine + BarCodeHeightX20 + BarCodeBelow + BarCodeFormB + barcode + NewLine;
        }

        public void BarCode(string bardcode, int model)
        {
            if (model == 765 || model == 605)
                BarCodeModelA(bardcode);
            else
                BarCodeModelB(bardcode);
        }

        public void Print()
        {
            string text = Initialize + FontCodePage0 + "RawPrinterUnicode" + NewLine + "CodePage 0"+ NewLine + "Normalize"  + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage2 + "RawPrinterUnicode" + NewLine + "CodePage 2" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage3 + "RawPrinterUnicode" + NewLine + "CodePage 3" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage4 + "RawPrinterUnicode" + NewLine + "CodePage 4" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage5 + "RawPrinterUnicode" + NewLine + "CodePage 5" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage17 + "RawPrinterUnicode" + NewLine + "CodePage 17" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage18 + "RawPrinterUnicode" + NewLine + "CodePage 18" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage19 + "RawPrinterUnicode" + NewLine + "CodePage 19" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage255 + "RawPrinterUnicode" + NewLine + "CodePage 255" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            text = InitializeFull + "RawPrinterUnicode" + NewLine + "FULL" + NewLine + "Normalize" + NewLine + Text + CutPaper;
            text = text.Normalize();
            SendStringToPrinter(PrinterName, text);

            // ------------------------------------------------------------

            text = Initialize + FontCodePage0 + "RawPrinterUnicode" + NewLine + "CodePage 0" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage2 + "RawPrinterUnicode" + NewLine + "CodePage 2" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage3 + "RawPrinterUnicode" + NewLine + "CodePage 3" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage4 + "RawPrinterUnicode" + NewLine + "CodePage 4" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage5 + "RawPrinterUnicode" + NewLine + "CodePage 5" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage17 + "RawPrinterUnicode" + NewLine + "CodePage 17" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage18 + "RawPrinterUnicode" + NewLine + "CodePage 18" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage19 + "RawPrinterUnicode" + NewLine + "CodePage 19" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage255 + "RawPrinterUnicode" + NewLine + "CodePage 255" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            text = InitializeFull + "RawPrinterUnicode" + NewLine + "FULL" + NewLine + "ASCII" + NewLine + Text + CutPaper;
            text = NormalizeCharactersAscii(text);
            SendStringToPrinter(PrinterName, text);

            // -----------------------------------------------------------------

            text = Initialize + FontCodePage0 + "RawPrinterUnicode" + NewLine + "CodePage 0" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage2 + "RawPrinterUnicode" + NewLine + "CodePage 2" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage3 + "RawPrinterUnicode" + NewLine + "CodePage 3" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage4 + "RawPrinterUnicode" + NewLine + "CodePage 4" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage5 + "RawPrinterUnicode" + NewLine + "CodePage 5" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage17 + "RawPrinterUnicode" + NewLine + "CodePage 17" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage18 + "RawPrinterUnicode" + NewLine + "CodePage 18" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage19 + "RawPrinterUnicode" + NewLine + "CodePage 19" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage255 + "RawPrinterUnicode" + NewLine + "CodePage 255" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            text = InitializeFull + "RawPrinterUnicode" + NewLine + "FULL" + NewLine + "UTF-8" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUtf8(text);
            SendStringToPrinter(PrinterName, text);

            // ------------------------------------------------

            text = Initialize + FontCodePage0 + "RawPrinterUnicode" + NewLine + "CodePage 0" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage2 + "RawPrinterUnicode" + NewLine + "CodePage 2" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage3 + "RawPrinterUnicode" + NewLine + "CodePage 3" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage4 + "RawPrinterUnicode" + NewLine + "CodePage 4" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage5 + "RawPrinterUnicode" + NewLine + "CodePage 5" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage17 + "RawPrinterUnicode" + NewLine + "CodePage 17" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage18 + "RawPrinterUnicode" + NewLine + "CodePage 18" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage19 + "RawPrinterUnicode" + NewLine + "CodePage 19" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = Initialize + FontCodePage255 + "RawPrinterUnicode" + NewLine + "CodePage 255" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);

            text = InitializeFull + "RawPrinterUnicode" + NewLine + "FULL" + NewLine + "UNICODE" + NewLine + Text + CutPaper;
            text = NormalizeCharactersUnicode(text);
            SendStringToPrinter(PrinterName, text);
        }

        public enum Align
        {
            Center,
            Left,
            Right
        }
    }

    class Program

    {
        static void Main(string[] args)
        {
            try
            {                                                
                const string asterics = "*****************************************************";
                const string line = "_____________________________________________________";
                const string please = "Por favor espere su turno";
                const string lblNumber = "SU NÚMERO ES";
                const string lblDate = "Fecha:";
                                
                Console.WriteLine(@"Se imprime la frase Á É Í Ó Ú. Probando 80 combinaciones de configuración.");
                Console.WriteLine(@"Imprimiendo ...");

                var turno = Turno.GetMock();
                var printerAnsi = new RawPrinterAnsi();

                printerAnsi.SetAlign(RawPrinterAnsi.Align.Center);
                printerAnsi.DrawLine(RawPrinterAnsi.SizeH1 + asterics);
                printerAnsi.DrawLine(RawPrinterAnsi.SizeH2 + turno.IdTurno);
                printerAnsi.DrawLine(RawPrinterAnsi.SizeH2 + lblNumber);
                printerAnsi.DrawLine(RawPrinterAnsi.SizeH4 + turno.Letra + turno.Numero);
                printerAnsi.DrawLine(RawPrinterAnsi.SizeH4 + turno.Letra + turno.Numero);
                printerAnsi.DrawLine(asterics);
                printerAnsi.DrawLine(RawPrinterAnsi.SizeH1 + asterics);
                printerAnsi.DrawLine(RawPrinterAnsi.SizeH1 + please);
                printerAnsi.DrawLine(RawPrinterAnsi.SizeH1 + lblDate + " " + turno.Fecha);
                printerAnsi.DrawLine(RawPrinterAnsi.SizeH1 + line);                
                foreach (var txt in turno.texts)
                {
                    printerAnsi.DrawLine(RawPrinterAnsi.SizeH1 + txt);
                }

                //var ascii = printer.NormalizeCharacters("ÁÉÍÓÚ");
                printerAnsi.BarCodeModelA("14159265");
                printerAnsi.BarCodeModelB("14159265");

                // Borramos, sólo imprmimos Ú.
                printerAnsi.Clear();
                printerAnsi.DrawLine("Á É Í Ó Ú");
                printerAnsi.Print();

                var printerUnicode = new RawPrinterAnsi();
                printerUnicode.DrawLine("Á É Í Ó Ú");
                printerUnicode.Print();

                Console.WriteLine(@"Impresión finaliza con éxito.");
                Console.WriteLine(@"Presione una tecla...");
                Console.ReadKey();                
            }
            catch (Exception ex)
            {
                Console.WriteLine(@"Impresión se canceló.");
                Console.WriteLine(ex.Message);
            }
        }        
    }

    public class Turno
    {
        public string IdTurno { get; set; }
        public string Numero { get; set; }
        public string Fecha { get; set; }
        public string Letra { get; set; }
        public string Tipo { get; set; }
        public List<string> texts { get; set; }

        public Turno()
        {
            texts = new List<string>();
        }

        public static Turno GetMock()
        {
            return new Turno
            {
                IdTurno = "1",
                Fecha = "01/01/2017",
                Letra = "A",
                Numero = "001",
                Tipo = "Vacunas",
                texts = new string[] { "text1", "text2", "text3" }.ToList()
            };
        }
    }

}
