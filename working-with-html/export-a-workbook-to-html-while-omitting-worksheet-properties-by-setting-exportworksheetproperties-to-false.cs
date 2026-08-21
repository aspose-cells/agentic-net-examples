// Title: Export Workbook to HTML without Worksheet Properties using Aspose.Cells (C#)
// Description: Learn how to save an Aspose.Cells workbook as HTML while suppressing worksheet property metadata by setting HtmlSaveOptions.ExportWorksheetProperties to false. The example creates a simple workbook, configures the option, and generates clean HTML output.
// Keywords: Aspose.Cells HTML export C# | ExportWorksheetProperties false | remove worksheet metadata HTML | save Excel as HTML Aspose | clean HTML report from Excel | global | USA | India
// Common Searches: Aspose.Cells disable worksheet properties when exporting to HTML | C# HtmlSaveOptions ExportWorksheetProperties example | How to generate HTML from Excel without sheet metadata | Aspose.Cells HTML export without properties tutorial | Export Excel to lightweight HTML using Aspose
// Developer Intent: Generate an HTML file from a workbook while omitting all worksheet property information.
// Use Cases: Create web‑ready reports from Excel data without extra metadata. | Produce compact HTML emails or newsletters that exclude sheet properties. | Batch‑convert multiple workbooks to HTML for archival storage, reducing file size.
// AI Prompts: Show a C# snippet that exports an Aspose.Cells workbook to HTML with ExportWorksheetProperties set to false and adds a custom stylesheet. | Explain how ExportWorksheetProperties influences the HTML output and how to verify its omission. | Provide a script that loops through all worksheets in a workbook, saving each as a separate HTML file while disabling worksheet properties.

using System;
using Aspose.Cells;

namespace ExportHtmlWithoutWorksheetProperties
{
    // Learn how to save an Aspose.Cells workbook as HTML while suppressing worksheet property metadata by setting HtmlSaveOptions.ExportWorksheetProperties to false. The example creates a simple workbook, configures the option, and generates clean HTML output.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // Configure HTML save options to omit worksheet properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportWorksheetProperties = false; // Disable worksheet properties export

            // Save the workbook as HTML
            string outputPath = "WorkbookWithoutWorksheetProps.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
