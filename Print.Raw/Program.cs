using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Print.Raw
{    
    public class RawPrinterHelper
    {
        #region Structure and API declarions ...
        
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
        }

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartDocPrinter(IntPtr hPrinter, int level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int size);

        #endregion

        // SendBytesToPrinter()
        // When the function is given a printer name and an unmanaged array
        // of bytes, the function sends those bytes to the print queue.
        // Returns true on success, false on failure.
        public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, int dwCount)
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

        public static bool SendFileToPrinter(string szPrinterName, string szFileName)
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

        public static bool SendStringToPrinter(string szPrinterName, string szString)
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

        public static string GetDefaultPrinter()
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

                string bold = Encoding.ASCII.GetString(new byte[] { 27, (byte)'|', (byte)'b', (byte)'C' });
                string underline = Encoding.ASCII.GetString(new byte[] { 27, (byte)'|', (byte)'2', (byte)'u', (byte)'C' });
                string italic = Encoding.ASCII.GetString(new byte[] { 27, (byte)'|', (byte)'i', (byte)'C' });
                string centerAlign = Encoding.ASCII.GetString(new byte[] { 27, (byte)'|', (byte)'c', (byte)'A' });
                string rightAlign = Encoding.ASCII.GetString(new byte[] { 27, (byte)'|', (byte)'r', (byte)'A' });
                string doubleWideCharacters = Encoding.ASCII.GetString(new byte[] { 27, (byte)'|', (byte)'2', (byte)'C' });
                string doubleHightCharacters = Encoding.ASCII.GetString(new byte[] { 27, (byte)'|', (byte)'3', (byte)'C' });
                string doubleWideAndHightCharacters = Encoding.ASCII.GetString(new byte[] { 27, (byte)'|', (byte)'4', (byte)'C' });

                const string sharpNewLine = "\n\r";
                const string sharpCutPapper = "\f";

                var turno = Turno.GetMock();
                var printerName = RawPrinterHelper.GetDefaultPrinter();

                Console.WriteLine(@"JUANMA se imprime un TURNO de prueba.");
                Console.WriteLine(@"Imprimiendo C# Nativo ...");

                var str = asterics + sharpNewLine;
                str += turno.IdTurno + sharpNewLine;
                str += lblNumber + sharpNewLine;
                str += turno.Letra + turno.Numero + sharpNewLine;
                str += turno.IdTurno + sharpNewLine;
                str += asterics + sharpNewLine;
                str += please + sharpNewLine;
                str += lblDate + " " + turno.Fecha + sharpNewLine;
                str += line + sharpNewLine;
                foreach (var txt in turno.Textos)
                {
                    str += txt + sharpNewLine;
                }
                str += sharpCutPapper;
                RawPrinterHelper.SendStringToPrinter(printerName, str);                

                Console.WriteLine(@"Impresión finaliza con éxito.");
                Console.WriteLine(@"Presione una tecla...");
                Console.ReadKey();

                Console.WriteLine(@"Imprimiendo con variables de PHP para cortar papel ...");
                string phpCupper = Convert.ToString((char)29) + Convert.ToString((char)105);

                str = string.Empty;
                str = asterics + sharpNewLine;
                str += turno.IdTurno + sharpNewLine;
                str += lblNumber + sharpNewLine;
                str += turno.Letra + turno.Numero + sharpNewLine;
                str += turno.IdTurno + sharpNewLine;
                str += asterics + sharpNewLine;
                str += please + sharpNewLine;
                str += lblDate + " " + turno.Fecha + sharpNewLine;
                str += line + sharpNewLine;
                foreach (var txt in turno.Textos)
                {
                    str += txt + sharpNewLine;
                }
                str += sharpCutPapper;
                RawPrinterHelper.SendStringToPrinter(printerName, str);

                Console.WriteLine(@"Impresión finaliza con éxito.");
                Console.WriteLine(@"Presione una tecla...");
                Console.ReadKey();

                Console.WriteLine(@"Imprimiendo con formato C# Nativo ...");                                
                RawPrinterHelper.SendStringToPrinter(printerName, centerAlign + asterics + sharpNewLine);
                RawPrinterHelper.SendStringToPrinter(printerName, bold + turno.IdTurno + sharpNewLine);               
                RawPrinterHelper.SendStringToPrinter(printerName, lblNumber + sharpNewLine);
                RawPrinterHelper.SendStringToPrinter(printerName, doubleWideAndHightCharacters + turno.Letra + turno.Numero + sharpNewLine);
                RawPrinterHelper.SendStringToPrinter(printerName, turno.IdTurno + sharpNewLine);
                RawPrinterHelper.SendStringToPrinter(printerName, asterics + sharpNewLine);
                RawPrinterHelper.SendStringToPrinter(printerName, please + sharpNewLine);
                RawPrinterHelper.SendStringToPrinter(printerName, lblDate + " " + turno.Fecha + sharpNewLine);
                RawPrinterHelper.SendStringToPrinter(printerName, line + sharpNewLine);
                foreach (var txt in turno.Textos)
                {
                    RawPrinterHelper.SendStringToPrinter(printerName, txt + sharpNewLine);
                }
                RawPrinterHelper.SendStringToPrinter(printerName,sharpCutPapper);                

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
        public List<string> Textos { get; set; }

        public Turno()
        {
            Textos = new List<string>();
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
                Textos = new string[] { "text1", "text2", "text3" }.ToList()
            };
        }
    }

}
