// Title: Batch convert Excel workbooks to HTML with conditional hidden‑sheet exclusion using Aspose.Cells for .NET
// Description: Iterates through a folder of .xlsx files, detects worksheets whose name contains "confidential" (case‑insensitive), and saves each workbook as HTML. The HtmlSaveOptions.ExportHiddenWorksheet flag is set to false only for workbooks that contain a confidential sheet, otherwise hidden sheets are included.
// Keywords: Aspose.Cells HTML export | ExportHiddenWorksheet conditional | batch Excel to HTML .NET | skip hidden worksheets confidential | C# Aspose.Cells example | process multiple workbooks | global Excel conversion | regional data privacy export
// Common Searches: Aspose.Cells export hidden worksheets based on sheet name | batch convert xlsx to html and hide confidential sheets | C# conditional ExportHiddenWorksheet option | how to exclude hidden sheets when exporting Excel to HTML | automate Excel to HTML conversion with privacy rules
// Developer Intent: Convert a collection of Excel files to HTML, disabling hidden‑sheet export only for workbooks that contain a worksheet named "confidential".
// Use Cases: Publish public HTML reports from a shared drive while automatically protecting hidden confidential data. | Nightly automation that transforms financial models to web‑ready format, preserving helper sheets unless a confidential tab is present. | Client‑side conversion service that delivers HTML versions of uploaded workbooks, omitting hidden sheets for privacy‑sensitive files.
// AI Prompts: Generate C# code with Aspose.Cells that batch converts .xlsx files to HTML, turning off ExportHiddenWorksheet when any worksheet name includes "confidential". | Add robust logging and error handling to the batch export program, recording files where hidden worksheets were excluded. | Create unit tests that verify ExportHiddenWorksheet is false only when a confidential worksheet exists in the workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExportHiddenWorksheets
{
    // Iterates through a folder of .xlsx files, detects worksheets whose name contains "confidential" (case‑insensitive), and saves each workbook as HTML. The HtmlSaveOptions.ExportHiddenWorksheet flag is set to false only for workbooks that contain a confidential sheet, otherwise hidden sheets are included.
    class Program
    {
        static void Main()
        {
            // Folder containing source Excel workbooks
            string sourceFolder = @"C:\InputWorkbooks";

            // Folder where the HTML files will be saved
            string outputFolder = @"C:\ExportedHtml";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the source folder
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Determine if any worksheet name contains the word "confidential" (case‑insensitive)
                bool containsConfidential = false;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.IndexOf("confidential", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        containsConfidential = true;
                        break;
                    }
                }

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Export hidden worksheets only when the workbook does NOT contain confidential sheets
                    ExportHiddenWorksheet = !containsConfidential
                };

                // Build the output HTML file path
                string outputFile = Path.Combine(
                    outputFolder,
                    Path.GetFileNameWithoutExtension(filePath) + ".html");

                // Save the workbook as HTML using the configured options
                workbook.Save(outputFile, saveOptions);
            }
        }
    }
}
