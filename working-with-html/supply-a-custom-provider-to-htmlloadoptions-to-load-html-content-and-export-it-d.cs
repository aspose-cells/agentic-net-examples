using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlStreamExample
{
    public class Program
    {
        public static void Main()
        {
            string html = @"
                <html>
                <body>
                    <table>
                        <tr><td>A1</td><td>123</td></tr>
                        <tr><td>A2</td><td>456</td></tr>
                    </table>
                </body>
                </html>";

            // Convert HTML string to a memory stream
            using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
            {
                // Set up load options for HTML
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();

                // Load the HTML into a workbook from the stream
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Export the loaded workbook to an XLSX file
                workbook.Save("ExportedFromHtml.xlsx");
            }
        }
    }
}