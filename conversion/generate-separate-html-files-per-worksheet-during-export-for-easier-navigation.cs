// Title: Export Each Excel Worksheet to Separate HTML Files with Aspose.Cells for .NET
// Description: Demonstrates how to iterate through a workbook, set each sheet as active, and save it as an individual HTML page using Aspose.Cells. The example skips evaluation‑warning sheets, creates file‑system‑safe names, and isolates CSS for faster loading.
// Keywords: Aspose.Cells HTML export | C# export worksheet to HTML | separate HTML per Excel sheet | ExportActiveWorksheetOnly | ExportWorksheetCSSSeparately | skip evaluation warning sheet | safe file name Excel | Aspose.Cells .NET tutorial
// Common Searches: how to save each Excel sheet as its own HTML file using Aspose.Cells | Aspose.Cells export active worksheet only | C# generate separate HTML pages for workbook worksheets | avoid evaluation warning sheets when exporting with Aspose.Cells | create file‑safe names from worksheet titles
// Developer Intent: Generate an individual HTML file for every worksheet in a workbook.
// Use Cases: Publish each department's data sheet as a standalone web page. | Create modular HTML reports that load CSS only for the relevant sheet. | Provide intranet users with direct links to specific worksheet content without navigating a large combined file.
// AI Prompts: Show how to attach a custom CSS stylesheet to each exported HTML file. | Convert each worksheet to a separate PDF using Aspose.Cells and preserve page layout. | Explain best practices for handling special characters in worksheet names when creating file names.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSeparateHtmlExport
{
    // Demonstrates how to iterate through a workbook, set each sheet as active, and save it as an individual HTML page using Aspose.Cells. The example skips evaluation‑warning sheets, creates file‑system‑safe names, and isolates CSS for faster loading.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Populate sample worksheets
                workbook.Worksheets[0].Name = "Summary";
                workbook.Worksheets[0].Cells["A1"].PutValue("Summary Data");

                Worksheet sheet1 = workbook.Worksheets.Add("Sales");
                sheet1.Cells["A1"].PutValue("Product");
                sheet1.Cells["B1"].PutValue("Quantity");
                sheet1.Cells["A2"].PutValue("Apple");
                sheet1.Cells["B2"].PutValue(150);

                Worksheet sheet2 = workbook.Worksheets.Add("Inventory");
                sheet2.Cells["A1"].PutValue("Item");
                sheet2.Cells["B1"].PutValue("Stock");
                sheet2.Cells["A2"].PutValue("Banana");
                sheet2.Cells["B2"].PutValue(80);

                // Prepare HTML save options to export only the active worksheet
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = true,
                    ExportWorksheetCSSSeparately = true
                };

                // Directory to store the generated HTML files
                string outputDir = Path.Combine(Environment.CurrentDirectory, "HtmlSheets");
                Directory.CreateDirectory(outputDir);

                // Export each worksheet (excluding evaluation warning sheets) to its own HTML file
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet ws = workbook.Worksheets[i];
                    string sheetName = ws.Name;

                    // Skip evaluation warning worksheets that Aspose adds in evaluation mode
                    if (sheetName.StartsWith("Evaluation Warning", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Set the current worksheet as active
                    workbook.Worksheets.ActiveSheetIndex = i;

                    // Build a safe file name based on the worksheet name
                    string safeName = string.Concat(sheetName.Split(Path.GetInvalidFileNameChars()));
                    string htmlPath = Path.Combine(outputDir, $"{safeName}.html");

                    // Save the active worksheet as HTML
                    workbook.Save(htmlPath, htmlOptions);
                    Console.WriteLine($"Exported worksheet '{sheetName}' to '{htmlPath}'.");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
