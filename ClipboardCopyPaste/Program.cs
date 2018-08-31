using System;
using System.Collections.Generic;
using System.IO;
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
            //// Excel Spreadsheet example
            //GetExcelSpreadsheet();
            //PasteExcelSpreadsheet();


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

        private static void GetExcelSpreadsheet()
        {
            var content = Clipboard.GetData("XML Spreadsheet");
            if (content is MemoryStream)
            {
                string decoded = Encoding.UTF8.GetString((content as MemoryStream).ToArray());
                Console.WriteLine(decoded);
            }
            else
            {
                Console.WriteLine(content);
            }
        }

        private static void PasteExcelSpreadsheet()
        {
            string content = @"<?xml version=""1.0""?>
<?mso-application progid=""Excel.Sheet""?>
<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet""
 xmlns:o=""urn:schemas-microsoft-com:office:office""
 xmlns:x=""urn:schemas-microsoft-com:office:excel""
 xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet""
 xmlns:html=""http://www.w3.org/TR/REC-html40"">
 <Styles>
  <Style ss:ID=""Default"" ss:Name=""Normal"">
   <Alignment ss:Vertical=""Bottom""/>
   <Borders/>
   <Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Size=""11"" ss:Color=""#000000""/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID=""s18"" ss:Name=""Currency"">
   <NumberFormat
    ss:Format=""_(&quot;$&quot;* #,##0.00_);_(&quot;$&quot;* \(#,##0.00\);_(&quot;$&quot;* &quot;-&quot;??_);_(@_)""/>
  </Style>
  <Style ss:ID=""s64"">
   <Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Size=""11"" ss:Color=""#000000""
    ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""s65"" ss:Parent=""s18"">
   <Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Size=""11"" ss:Color=""#000000""/>
  </Style>
 </Styles>
 <Worksheet ss:Name=""Sheet1"">
  <Table ss:ExpandedColumnCount=""2"" ss:ExpandedRowCount=""2""
   ss:DefaultRowHeight=""15"">
   <Row>
    <Cell><Data ss:Type=""String"">Month</Data></Cell>
    <Cell><Data ss:Type=""String"">Year</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID=""s64""><Data ss:Type=""String"">August</Data></Cell>
    <Cell ss:StyleID=""s65""><Data ss:Type=""Number"">999.99</Data></Cell>
   </Row>
  </Table>
 </Worksheet>
</Workbook>";

            // This is the worng way to copy a String value
            //Clipboard.SetData("XML Spreadsheet", content);

            // This is the right way to copyu a string value without .NET wrapping
            Clipboard.SetData("XML Spreadsheet", new MemoryStream(Encoding.UTF8.GetBytes(content)));
        }
    }
}
