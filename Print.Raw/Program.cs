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
        // Structure and API declarions:
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
        public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

        // SendBytesToPrinter()
        // When the function is given a printer name and an unmanaged array
        // of bytes, the function sends those bytes to the print queue.
        // Returns true on success, false on failure.
        public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
        {
            Int32 dwError = 0, dwWritten = 0;
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
            Byte[] bytes = new Byte[fs.Length];
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
            Int32 dwCount;

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
    }

    class Program

    {
        static void Main(string[] args)

        {

            try

            {

                var printer = new PrintDocument();
                printer.PrintPage += (sender, e) => PrintPage(sender, e, Turno.GetMock());
                printer.Print();

                Console.WriteLine(@"JUANMA se imprime el archivo C:\Sisfarma.es\TURNO.pdf");
                Console.WriteLine(@"Imprimiendo ...");

                RawPrinterHelper.SendStringToPrinter("Microsoft XPS Document Writer", "Hola mundo!");
                Console.WriteLine(@"Impresión finaliza con éxito.");
            }

            catch (Exception ex)

            {

                Console.WriteLine(@"Impresión se canceló.");
                Console.WriteLine(ex.Message);

            }

        }

        private static void PrintPage(object sender, PrintPageEventArgs e, Turno turno)
        {
            var currentY = 0f;
            var text = "*****************************************************";
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);

            text = turno.IdTurno;
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 20f, FontStyle.Bold), StringAlignment.Center, ref currentY);

            text = "SU NÚMERO ES";
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 15f, FontStyle.Bold), StringAlignment.Center, ref currentY);

            text = turno.Letra + turno.Numero;
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 50f, FontStyle.Bold), StringAlignment.Center, ref currentY);

            text = "*****************************************************";
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);

            text = "Por favor espere su turno";
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);

            text = "Fecha: " + turno.Fecha;
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);

            text = "_____________________________________________________";
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);

            foreach (var item in turno.Textos)
            {
                PrintLine(item, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);
            }

            e.HasMorePages = false;
        }

        private static void PrintLine(string text, PrintPageEventArgs e, Font font, StringAlignment alignment, ref float currentY)
        {
            using (var sf = new StringFormat())
            {
                sf.Alignment = alignment;
                var size = e.Graphics.MeasureString(text, font, e.MarginBounds.Size);
                e.Graphics.DrawString(text, font, Brushes.Black, new PointF(e.MarginBounds.Left + e.MarginBounds.Width / 2, currentY), sf);
                currentY += font.GetHeight(e.Graphics);
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
