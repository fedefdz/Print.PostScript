using System;
using System.Windows.Forms; // required for Form class
using System.Drawing; // required for Bitmap class
using System.Drawing.Printing; // required for printPreviewDocument class

namespace Print.Direct.Bit
{


    // ICE ROMEO
    // ibknowbodi@yahoo.com
    // CLASS FOR PRINTING MADE EASY    

    namespace PrintFormToPrinter

    {

        /// <summary>
        /// FormPrinter is a class that I created to give
        /// users a simple way to print forms. The FormPrinter
        /// class receives a form in its constructor and
        /// prints that form to the default printer (which
        /// can be changed). This code is based on code found
        /// in the .NET help system.
        /// Rick Bird, DeVry University, June 6, 2004
        /// </summary>

        public class FormPrinter

        {
            // attributes
            private Form formToPrint;
            // create necessary printing objects
            PrintDocument printDoc = new PrintDocument(); // required PrintDocument object
            PrintDialog printDialog = new PrintDialog(); // required PrintDialog object
            PrintPreviewDialog printPreviewDialog = new PrintPreviewDialog(); // required PrintPreviewDialog object

            // bring in the BitBlt windows API method and necessary references
            [System.Runtime.InteropServices.DllImport("gdi32.dll")]
            public static extern long BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, int dwRop);

            private Bitmap memoryImage;
            
            // constructors
            public FormPrinter()

            {
                formToPrint = null;
                // create an PrintPage event for the printPreviewDoc object
                this.printDoc.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDoc_PrintPage);
            }

            public FormPrinter(Form formToPrint)
            {
                this.formToPrint = formToPrint;
                // create an PrintPage event for the printPreviewDoc object
                this.printDoc.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDoc_PrintPage);

            }

            // behaviors

            public void printWithDialog()

            {
                bool formWasOpened = false;
                // set the printer defaults
                bool printerSet = setPrinterDefaults();
                printPreviewDialog.Document = printDoc;
                // check to see if the form is open
                if (formToPrint.Visible == false)
                {
                    formToPrint.Show();
                    formWasOpened = true;
                }

                formToPrint.Refresh();

                // print the form unless user hits cancel
                if (printerSet == true)

                {
                    captureScreen();
                    // use printPreviewDialog.ShowDialog() to show a preview dialog of form
                    //printPreviewDialog.ShowDialog();
                    // use printDoc.Print() to print immediately after choosing printer
                    printDoc.Print();
                }

                // if we had to open the form, close it now

                if (formWasOpened == true)
                {
                    formToPrint.Dispose();
                }

            }

            public void printWithOutDialog(Form formToPrint)

            {

                // set form to object level
                this.formToPrint = formToPrint;
                bool formWasOpened = false;
                // set the printer defaults
                printPreviewDialog.Document = printDoc;
                // check to see if the form is open

                if (formToPrint.Visible == false)

                {
                    formToPrint.Show();
                    formWasOpened = true;
                }

                formToPrint.Refresh();
                formToPrint.Refresh();
                // print the form unless user hits cancel
                captureScreen();
                // use printPreviewDialog.ShowDialog() to show a preview dialog of form
                //printPreviewDialog.ShowDialog();
                // use printDoc.Print() to print immediately after choosing printer
                printDoc.Print();

                // if we had to open the form, close it now
                if (formWasOpened == true)
                {
                    formToPrint.Dispose();
                }

            }

            private bool setPrinterDefaults()

            {

                printDialog.Document = printDoc;
                DialogResult dialogRslt = printDialog.ShowDialog();
                if (dialogRslt == DialogResult.OK)
                {

                    printDoc.PrinterSettings = printDialog.PrinterSettings;
                    // set printer to high resolution
                    printDoc.DefaultPageSettings.PrinterResolution =
                    printDoc.PrinterSettings.PrinterResolutions[0]; // 0=high, 1=med, 2=low, 3=draft (may change depending on your printer)
                    return true;
                }

                return false; // return false if user cancels printer dialog

            }

            private void captureScreen()

            {
                // get the current screen
                Graphics mygraphics = formToPrint.CreateGraphics();
                Size s = formToPrint.Size;
                memoryImage = new Bitmap(s.Width, s.Height, mygraphics);
                Graphics memoryGraphics = Graphics.FromImage(memoryImage);
                IntPtr dc1 = mygraphics.GetHdc();
                IntPtr dc2 = memoryGraphics.GetHdc();
                BitBlt(dc2, 0, 0, formToPrint.ClientRectangle.Width, formToPrint.ClientRectangle.Height, dc1, 0, 0, 13369376);
                mygraphics.ReleaseHdc(dc1);
                memoryGraphics.ReleaseHdc(dc2);
            }

            private void printDoc_PrintPage(System.Object sender, System.Drawing.Printing.PrintPageEventArgs e)

            {
                // this event fires when page is printing
                e.Graphics.DrawImage(memoryImage, 0, 0);
            }

            // accessors and modifiers
            public Form getFormToPrint()
            {
                return formToPrint;
            }

            public void setFormToPrint(Form formToPrint)

            {
                this.formToPrint = formToPrint;
            }
        }
    }
}
