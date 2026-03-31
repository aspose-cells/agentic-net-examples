using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlImport
{
    class Program
    {
        static void Main()
        {
            // Sample HTML content containing a self‑closing tag (<br/>)
            string htmlContent = @"
                <html>
                    <body>
                        <p>First line<br/>Second line</p>
                    </body>
                </html>";

            // Convert the HTML string to a UTF‑8 encoded memory stream
            using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                // Create HtmlLoadOptions and enable support for self‑closing tags
                // (SupportDivTag also enables proper handling of layout tags; self‑closing tags are recognized automatically)
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                loadOptions.SupportDivTag = true;          // enable layout handling
                loadOptions.AutoFitColsAndRows = true;    // optional: auto‑fit after import

                // Load the HTML content into a Workbook using the specified options
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Export the workbook to an XLSX file
                workbook.Save("ImportedFromHtml.xlsx");
            }

            Console.WriteLine("HTML content has been successfully imported and saved as XLSX.");
        }
    }
}