// Title: Export AspNet Aspose.Cells Workbook to HTML using a MemoryStream (C#)
// Description: Demonstrates how to generate HTML from an Aspose.Cells workbook directly into a MemoryStream, embed images as Base64, and read the result as a string—eliminating the need for a temporary file on disk.
// Keywords: Aspose.Cells HTML MemoryStream | C# export workbook to HTML in memory | embed images Base64 Aspose.Cells | HtmlSaveOptions ExportActiveWorksheetOnly | CustomImplementationFactory MemoryStream | Aspose.Cells in‑memory HTML conversion
// Common Searches: Aspose.Cells save workbook as HTML to MemoryStream C# | export Excel to HTML without creating a file | generate HTML from spreadsheet in memory Aspose | read HTML output from Aspose.Cells MemoryStream
// Developer Intent: Generate an HTML representation of a spreadsheet and keep it in RAM for immediate consumption, such as sending it over a network or embedding it in another document.
// Use Cases: Return spreadsheet HTML from a Web API endpoint without writing to disk. | Compose an email body that contains the workbook data as HTML. | Apply string manipulation or templating to the HTML before storing or transmitting it.
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook to a MemoryStream as HTML with Base64‑encoded images. | Show how to modify the sample to export all worksheets instead of only the active one. | Explain how to stream the generated HTML directly as the response body in an ASP.NET Core controller.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsInMemoryHtmlExport
{
    // Demonstrates how to generate HTML from an Aspose.Cells workbook directly into a MemoryStream, embed images as Base64, and read the result as a string—eliminating the need for a temporary file on disk.
    class Program
    {
        static void Main()
        {
            // 1. Create a sample workbook and add some data.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells HTML Export to MemoryStream");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(12345);
            sheet.Cells["B2"].PutValue(3.14159);

            // 2. Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Embed images as Base64 to avoid external resource files.
            htmlOptions.ExportImagesAsBase64 = true;
            // Optional: export only the active worksheet.
            htmlOptions.ExportActiveWorksheetOnly = true;

            // 3. Create a MemoryStream using the provided CustomImplementationFactory rule.
            CustomImplementationFactory factory = new CustomImplementationFactory();
            MemoryStream htmlStream = factory.CreateMemoryStream();

            // 4. Save the workbook as HTML into the MemoryStream.
            workbook.Save(htmlStream, htmlOptions);

            // 5. Reset the stream position to read the generated HTML.
            htmlStream.Position = 0;
            string htmlContent = Encoding.UTF8.GetString(htmlStream.ToArray());

            // 6. Demonstrate that the HTML is available in memory.
            Console.WriteLine("Generated HTML content (first 200 characters):");
            Console.WriteLine(htmlContent.Substring(0, Math.Min(200, htmlContent.Length)));

            // Clean up.
            htmlStream.Dispose();
            workbook.Dispose();
        }
    }
}
