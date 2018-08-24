using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClipboardCopyPaste
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var currentData = Clipboard.GetDataObject();
            var newData = new System.Windows.Forms.DataObject();

            // Go through all the formats in the clipboard, and only copy
            // the formats selected with CopyDataFromCurrentToNew().
            // All other formats will be removed from the clipboard.
            foreach (string format in currentData.GetFormats())
            {
                // Sometimes CF_BITMAP show as a zero length content,
                // but Handle Type is bitmap, which apparently is treated differently.
                // Is there really an image in the clipboard?
                Console.WriteLine(format);
                if (format == "Bitmap")
                {
                    var bitmap = currentData.GetData(format) as System.Drawing.Bitmap;
                    Console.WriteLine("{0}x{1}", bitmap.Width, bitmap.Height);
                }

                //CopyDataFromCurrentToNew(currentData, newData, format, "CF_UNICODETEXT");
                CopyDataFromCurrentToNew(currentData, newData, format, "PNG");
            }

            Clipboard.SetDataObject(newData, true);
        }

        private static void CopyDataFromCurrentToNew(IDataObject currentData, DataObject newData, string currentFormat, string expectedFormat)
        {
            if (currentFormat == expectedFormat)
            {
                var formattedData = currentData.GetData(currentFormat);
                newData.SetData(currentFormat, formattedData);
            }
        }
    }
}
