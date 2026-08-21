// Title: Export Worksheet‑Specific CSS Files When Saving Excel to HTML with Aspose.Cells for .NET
// Description: Demonstrates how to set HtmlSaveOptions.ExportWorksheetCSSSeparately to true so each worksheet in a workbook is saved with its own CSS file, producing separate stylesheet files alongside the HTML output.
// Keywords: Aspose.Cells ExportWorksheetCSSSeparately | HTML save options per worksheet CSS | C# Aspose.Cells separate CSS files | Excel to HTML with individual stylesheets | .NET generate per‑sheet CSS
// Common Searches: Aspose.Cells ExportWorksheetCSSSeparately example C# | save Excel workbook as HTML with separate CSS per sheet | how to generate individual CSS files for each worksheet using Aspose.Cells | HTML export options for multi‑sheet workbook Aspose.Cells
// Developer Intent: Create distinct CSS files for each worksheet when converting an Excel workbook to HTML.
// Use Cases: Publish multi‑sheet Excel reports on the web where each sheet needs its own styling. | Maintain separate branding or themes for individual worksheets in HTML output. | Automate generation of HTML dashboards from Excel with per‑sheet CSS for easier updates.
// AI Prompts: Show C# code that saves an Excel workbook to HTML with ExportWorksheetCSSSeparately enabled using Aspose.Cells. | Explain the folder structure created when ExportWorksheetCSSSeparately is true and how to reference the generated CSS files. | Provide guidance on customizing the names of the CSS files produced by ExportWorksheetCSSSeparately in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to set HtmlSaveOptions.ExportWorksheetCSSSeparately to true so each worksheet in a workbook is saved with its own CSS file, producing separate stylesheet files alongside the HTML output.
    public class ExportWorksheetCssSeparatelyDemo
    {
        public static void Run()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue("Data in Sheet 1");

            // Add a second worksheet with its own data
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].PutValue("Data in Sheet 2");

            // Configure HTML save options to export CSS for each worksheet separately
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportWorksheetCSSSeparately = true // Enable separate CSS files per worksheet
            };

            // Define output directory and file name
            string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "HtmlExport");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "Workbook.html");

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
            Console.WriteLine("Separate CSS files for each worksheet are generated in the same directory.");
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
