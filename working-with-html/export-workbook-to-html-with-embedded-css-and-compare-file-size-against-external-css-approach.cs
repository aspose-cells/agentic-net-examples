// Title: Aspose.Cells C# Export Workbook to HTML – Embedded vs External CSS with File‑Size Comparison
// Description: Demonstrates how to save an Aspose.Cells workbook as HTML using the default embedded CSS and using external CSS via HtmlSaveOptions.ExportWorksheetCSSSeparately. The sample measures the HTML and CSS file sizes, prints a side‑by‑side comparison, and indicates which approach yields a smaller total payload.
// Keywords: Aspose.Cells HTML export | C# HtmlSaveOptions | ExportWorksheetCSSSeparately | embedded CSS vs external CSS | HTML file size comparison | Aspose.Cells CSS optimization | Excel to HTML conversion .NET
// Common Searches: Aspose.Cells export HTML with embedded CSS C# | How to generate external CSS file when saving Excel as HTML using Aspose.Cells | Compare HTML file size with embedded and external CSS in Aspose.Cells | HtmlSaveOptions ExportWorksheetCSSSeparately example | C# code to measure size of Aspose.Cells HTML output
// Developer Intent: Create HTML from a workbook with both embedded and external CSS, then evaluate which method produces a smaller overall file size.
// Use Cases: Determine the most efficient CSS strategy for small reports to minimize HTTP requests. | Generate reusable stylesheet files for large workbooks so browsers can cache the CSS across pages. | Integrate an automated size‑comparison step into a reporting pipeline to select the optimal export option.
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook to HTML with embedded CSS and outputs the file size. | Show how to configure HtmlSaveOptions to export CSS to a separate file and calculate the combined size of HTML and CSS. | Explain how to extend the example to log the size comparison results to a CSV file for later analysis.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
    // Demonstrates how to save an Aspose.Cells workbook as HTML using the default embedded CSS and using external CSS via HtmlSaveOptions.ExportWorksheetCSSSeparately. The sample measures the HTML and CSS file sizes, prints a side‑by‑side comparison, and indicates which approach yields a smaller total payload.
    class Program
    {
        static void Main()
        {
            // Create a sample workbook with styled data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with different styles to generate CSS rules
            sheet.Cells["A1"].PutValue("Header");
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = System.Drawing.Color.White;
            headerStyle.ForegroundColor = System.Drawing.Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["A1"].SetStyle(headerStyle);

            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["B3"].PutValue(456);

            // Apply a different style to the numeric column
            Style numberStyle = workbook.CreateStyle();
            numberStyle.Font.Color = System.Drawing.Color.Green;
            numberStyle.Number = 3; // "#,##0"
            sheet.Cells["B2"].SetStyle(numberStyle);
            sheet.Cells["B3"].SetStyle(numberStyle);

            // -----------------------------------------------------------------
            // 1) Export with embedded CSS (default behavior)
            // -----------------------------------------------------------------
            HtmlSaveOptions embeddedOptions = new HtmlSaveOptions();
            // ExportWorksheetCSSSeparately defaults to false, so CSS will be embedded
            string embeddedHtmlPath = "EmbeddedCssOutput.html";
            workbook.Save(embeddedHtmlPath, embeddedOptions);

            // Get size of the HTML file with embedded CSS
            long embeddedHtmlSize = new FileInfo(embeddedHtmlPath).Length;

            // -----------------------------------------------------------------
            // 2) Export with external CSS (ExportWorksheetCSSSeparately = true)
            // -----------------------------------------------------------------
            // Define a folder where the external CSS file will be written
            string externalFolder = "ExternalCssFiles";
            Directory.CreateDirectory(externalFolder);

            HtmlSaveOptions externalOptions = new HtmlSaveOptions();
            externalOptions.ExportWorksheetCSSSeparately = true;
            externalOptions.AttachedFilesDirectory = externalFolder; // folder for CSS file
            string externalHtmlPath = Path.Combine(externalFolder, "ExternalCssOutput.html");
            workbook.Save(externalHtmlPath, externalOptions);

            // Get size of the HTML file (without CSS) and the generated CSS file
            long externalHtmlSize = new FileInfo(externalHtmlPath).Length;

            // The CSS file name follows the pattern "sheet0.css"
            string externalCssPath = Path.Combine(externalFolder, "sheet0.css");
            long externalCssSize = File.Exists(externalCssPath) ? new FileInfo(externalCssPath).Length : 0;

            // -----------------------------------------------------------------
            // Output comparison results
            // -----------------------------------------------------------------
            Console.WriteLine("File size comparison:");
            Console.WriteLine($"Embedded CSS HTML size : {embeddedHtmlSize} bytes");
            Console.WriteLine($"External CSS HTML size : {externalHtmlSize} bytes");
            Console.WriteLine($"External CSS file size : {externalCssSize} bytes");
            Console.WriteLine($"Total size (HTML + CSS) : {externalHtmlSize + externalCssSize} bytes");

            // Simple decision output
            if (embeddedHtmlSize < externalHtmlSize + externalCssSize)
                Console.WriteLine("Embedded CSS approach results in a smaller overall file size.");
            else if (embeddedHtmlSize > externalHtmlSize + externalCssSize)
                Console.WriteLine("External CSS approach results in a smaller overall file size.");
            else
                Console.WriteLine("Both approaches produce the same total file size.");
        }
    }
}
