using System;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Print.PostScript;

namespace Print.PostScript
{
    public class RawPrinterHelper
    {
        // Estructturas y APIs:
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        private class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDataType;
        }

        #region dll Wrappers
        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int size);

        #endregion dll Wrappers

        #region Methods
        /// <summary>
        /// A partir de la ruta y nombre de un archivo pdf.
        /// Abre el archivo, obtenos sus btyes y envía en forma plana a imprimir.        
        /// </summary>
        /// <param name="szPrinterName">Nombre de impresora</param>
        /// <param name="szFileName">Ruta y nombre del archivo</param>
        /// <returns>true en caso de éxito, false de lo contrario</returns>
        public static bool SendFileToPrinter(string pdfFileName)
        {
            try
            {
                #region Get Connected Printer Name
                PrintDocument pd = new PrintDocument();
                StringBuilder dp = new StringBuilder(256);
                int size = dp.Capacity;
                if (GetDefaultPrinter(dp, ref size))
                {
                    pd.PrinterSettings.PrinterName = dp.ToString().Trim();
                }
                #endregion Get Connected Printer Name

                
                FileStream fs = new FileStream(pdfFileName, FileMode.Open);
                

                BinaryReader br = new BinaryReader(fs);
                Byte[] bytes = new Byte[fs.Length];
                bool success = false;
                
                
                // Unmanaged pointer.
                IntPtr ptrUnmanagedBytes = new IntPtr(0);
                int nLength = Convert.ToInt32(fs.Length);
                
                // Read contents of the file into the array.
                bytes = br.ReadBytes(nLength);
                // Allocate some unmanaged memory for those bytes.
                ptrUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
                // Copy the managed byte array into the unmanaged array.
                Marshal.Copy(bytes, 0, ptrUnmanagedBytes, nLength);
                // Send the unmanaged bytes to the printer.
                success = SendBytesToPrinter(pd.PrinterSettings.PrinterName, ptrUnmanagedBytes, nLength);
                // Free the unmanaged memory that you allocated earlier.
                Marshal.FreeCoTaskMem(ptrUnmanagedBytes);
                return success;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// This function gets the printer name and an unmanaged array of bytes, the function sends those bytes to the print queue.
        /// </summary>
        /// <param name="szPrinterName">Printer Name</param>
        /// <param name="pBytes">No. of bytes in the pdf file</param>
        /// <param name="dwCount">Word count</param>
        /// <returns>True on success, false on failure</returns>
        private static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
        {
            try
            {
                Int32 dwError = 0, dwWritten = 0;
                IntPtr hPrinter = new IntPtr(0);
                DOCINFOA di = new DOCINFOA();
                bool success = false; // Assume failure unless you specifically succeed.

                di.pDocName = "Ticket Document";
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
                            // Write the bytes.
                            success = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                            EndPagePrinter(hPrinter);
                        }
                        EndDocPrinter(hPrinter);
                    }
                    ClosePrinter(hPrinter);
                }

                // If print did not succeed, GetLastError may give more information about the failure.
                if (success == false)
                {
                    dwError = Marshal.GetLastWin32Error();
                }
                return success;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        #endregion Methods
    }
}
//Testing above code

class Program

{
    static void Main(string[] args)

    {

        try

        {
            Console.WriteLine(@"JUANMA se imprime el archivo C:\Sisfarma.es\TURNO.pdf");
            Console.WriteLine(@"Imprimiendo ...");

            RawPrinterHelper.SendFileToPrinter(@"C:\Sisfarma.es\TURNO.pdf");
            Console.WriteLine(@"Impresión finaliza con éxito.");
        }

        catch (Exception ex)

        {

            Console.WriteLine(@"Impresión se canceló.");
            Console.WriteLine(ex.Message);

        }

    }

}
