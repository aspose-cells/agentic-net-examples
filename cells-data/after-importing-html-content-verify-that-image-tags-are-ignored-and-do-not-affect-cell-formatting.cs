using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlImageIgnoreDemo
{
    class Program
    {
        static void Main()
        {
            // Sample HTML containing a table with an <img> tag inside a cell
            string htmlContent = @"
                <html>
                    <body>
                        <table>
                            <tr>
                                <td>Sample Text<img src='sample.png' alt='Image'></td>
                            </tr>
                        </table>
                    </body>
                </html>";

            // Convert the HTML string to a memory stream (required for loading with options)
            byte[] htmlBytes = System.Text.Encoding.UTF8.GetBytes(htmlContent);
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                // Create load options – default options are sufficient for this test
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();

                // Load the HTML into a workbook using the stream and load options
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Access the first worksheet and the first cell (A1) where the text resides
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A1"];

                // Verify that the cell value contains only the text part (image tag ignored)
                Console.WriteLine("Cell A1 value: " + cell.StringValue); // Expected: "Sample Text"

                // Verify that no picture objects were created from the <img> tag
                int pictureCount = sheet.Pictures.Count;
                Console.WriteLine("Number of picture objects in worksheet: " + pictureCount); // Expected: 0

                // Additionally, ensure the cell does not have an embedded image
                bool hasEmbeddedImage = cell.EmbeddedImage != null;
                Console.WriteLine("Cell A1 has embedded image: " + hasEmbeddedImage); // Expected: False

                // Save the workbook to verify that the import succeeded without images
                workbook.Save("HtmlImport_NoImages.xlsx", SaveFormat.Xlsx);
            }
        }
    }
}