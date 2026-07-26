// Title: Export Excel Workbook to HTML with Individual CSS Files per Worksheet using Aspose.Cells (C#)
// Description: Creates a three‑sheet workbook, enables HtmlSaveOptions.ExportWorksheetCSSSeparately, saves the workbook as HTML, and programmatically verifies that the number of generated *.css files matches the worksheet count, confirming a separate CSS file for each sheet.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportWorksheetCSSSeparately | C# | .NET | HTML export | separate CSS per worksheet | workbook to HTML | verify CSS files | multi‑sheet Excel | CSS isolation
// Common Searches: Aspose.Cells export each worksheet to its own CSS file | C# HtmlSaveOptions ExportWorksheetCSSSeparately example | how to verify CSS files after exporting Excel to HTML | generate separate CSS for Excel sheets with Aspose | HTML export of multi‑sheet workbook with individual styles
// Developer Intent: Generate HTML where every worksheet receives a dedicated CSS file and automatically confirm that the expected number of CSS files is created.
// Use Cases: Publish a multi‑sheet Excel report on a website with isolated styling for each sheet. | Add a CI validation step that checks the ExportWorksheetCSSSeparately option produces the correct CSS file count. | Build a web viewer that can apply distinct themes to individual worksheets by editing their separate CSS files.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as HTML with ExportWorksheetCSSSeparately enabled and store the CSS files in a target folder. | Create a method that scans a directory for *.css files, compares the count to the workbook's worksheet count, and logs a pass/fail result. | Provide a complete example that creates several worksheets, populates them, exports to HTML with separate CSS per sheet, and prints the names of the generated CSS files.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Creates a three‑sheet workbook, enables HtmlSaveOptions.ExportWorksheetCSSSeparately, saves the workbook as HTML, and programmatically verifies that the number of generated *.css files matches the worksheet count, confirming a separate CSS file for each sheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Populate each worksheet with some data
            workbook.Worksheets[0].Cells["A1"].PutValue("Data in Sheet1");
            workbook.Worksheets[1].Cells["A1"].PutValue("Data in Sheet2");
            workbook.Worksheets[2].Cells["A1"].PutValue("Data in Sheet3");

            // Configure HTML save options to export CSS separately for each worksheet
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportWorksheetCSSSeparately = true;

            // Define output directory and ensure it exists
            string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "HtmlExport");
            Directory.CreateDirectory(outputDir);

            // Save the workbook as HTML; each worksheet will generate its own .css file in the same folder
            string htmlPath = Path.Combine(outputDir, "Workbook.html");
            workbook.Save(htmlPath, saveOptions);

            // Verify that each worksheet has its own CSS file
            int worksheetCount = workbook.Worksheets.Count;
            string[] cssFiles = Directory.GetFiles(outputDir, "*.css");

            Console.WriteLine($"Total worksheets: {worksheetCount}");
            Console.WriteLine($"CSS files found: {cssFiles.Length}");

            // Simple verification: number of CSS files should match number of worksheets
            if (cssFiles.Length == worksheetCount)
            {
                Console.WriteLine("Verification passed: each worksheet has its own CSS file.");
                foreach (string cssFile in cssFiles)
                {
                    Console.WriteLine($" - {Path.GetFileName(cssFile)}");
                }
            }
            else
            {
                Console.WriteLine("Verification failed: mismatch between worksheets and CSS files.");
                Console.WriteLine("Expected CSS files:");
                for (int i = 0; i < worksheetCount; i++)
                {
                    Console.WriteLine($" - sheet{i + 1}.css (or similar)");
                }
                Console.WriteLine("Actual CSS files:");
                foreach (string cssFile in cssFiles)
                {
                    Console.WriteLine($" - {Path.GetFileName(cssFile)}");
                }
            }
        }
    }
}
