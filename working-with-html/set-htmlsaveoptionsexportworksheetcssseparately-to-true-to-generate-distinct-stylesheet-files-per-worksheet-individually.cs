// Title: Aspose.Cells .NET: Export Worksheet‑Specific CSS Files When Saving to HTML
// Description: Demonstrates how to set HtmlSaveOptions.ExportWorksheetCSSSeparately to true so that each worksheet in a workbook is saved with its own CSS stylesheet file alongside the generated HTML.
// Keywords: Aspose.Cells ExportWorksheetCSSSeparately | HTML save options .NET | per‑sheet CSS Aspose | generate separate CSS files | Workbook to HTML Aspose.Cells
// Common Searches: Aspose.Cells ExportWorksheetCSSSeparately example C# | save workbook as HTML with individual CSS files | how to create separate stylesheet per worksheet Aspose | HtmlSaveOptions per‑sheet CSS .NET
// Developer Intent: Save a workbook as HTML while producing an independent CSS file for each worksheet.
// Use Cases: Web reports that need isolated styling for each sheet to prevent CSS conflicts. | Performance‑optimized portals that cache per‑sheet styles separately. | Documentation generators that export multi‑sheet workbooks with distinct style sheets.
// AI Prompts: Generate code to customize the filenames of the CSS files created when ExportWorksheetCSSSeparately is enabled. | Show how to programmatically insert <link> tags for the separate CSS files into a custom HTML template after saving. | Explain the interaction between ExportWorksheetCSSSeparately and other HtmlSaveOptions such as ExportImagesAsBase64 and PreserveOriginalColumnWidth.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportWorksheetCssSeparatelyDemo
{
    // Demonstrates how to set HtmlSaveOptions.ExportWorksheetCSSSeparately to true so that each worksheet in a workbook is saved with its own CSS stylesheet file alongside the generated HTML.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data to two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Data in first sheet");

            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["A1"].PutValue("Data in second sheet");

            // Configure HTML save options to export CSS for each worksheet separately
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportWorksheetCSSSeparately = true; // Generates distinct CSS files per worksheet

            // Define the output directory and file name
            string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AsposeHtmlExport");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "Workbook.html");

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
            Console.WriteLine("Separate CSS files for each worksheet have been generated in the same directory.");
        }
    }
}
