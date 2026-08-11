// Title: Export Excel worksheets to HTML with individual external CSS files using Aspose.Cells for .NET
// Description: Learn how to save a workbook as HTML where each worksheet gets its own external CSS file. The example shows creating a workbook, setting HtmlSaveOptions.ExportWorksheetCSSSeparately = true, defining an assets folder with AttachedFilesDirectory, and generating HTML that correctly links to per‑sheet stylesheets.
// Keywords: Aspose.Cells HTML export | ExportWorksheetCSSSeparately | per‑worksheet CSS | HtmlSaveOptions AttachedFilesDirectory | C# Excel to HTML | external stylesheet Aspose | Aspose.Cells .NET example
// Common Searches: Aspose.Cells export each worksheet to separate CSS file | HTML export with per‑sheet stylesheet C# | HtmlSaveOptions ExportWorksheetCSSSeparately usage | Set folder for Aspose.Cells HTML assets | How to link external CSS for each Excel sheet in Aspose
// Developer Intent: Generate HTML from an Excel workbook where every worksheet references its own external CSS file located in a specified assets directory.
// Use Cases: Build modular web reports that load sheet‑specific styles from isolated CSS files. | Automate bulk conversion of Excel workbooks to HTML while keeping styling files organized per worksheet. | Integrate exported HTML into documentation portals that require separate stylesheet files for each page.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, enabling ExportWorksheetCSSSeparately and setting a custom assets folder. | Explain how HtmlSaveOptions.AttachedFilesDirectory determines the location of generated CSS files and how the main HTML file references them. | Provide troubleshooting steps when the exported HTML does not correctly link to the per‑worksheet CSS files.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates exporting each worksheet with its own external CSS file.
    // Learn how to save a workbook as HTML where each worksheet gets its own external CSS file. The example shows creating a workbook, setting HtmlSaveOptions.ExportWorksheetCSSSeparately = true, defining an assets folder with AttachedFilesDirectory, and generating HTML that correctly links to per‑sheet stylesheets.
    public class ExportWorksheetCssSeparatelyDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample data to two worksheets.
                Workbook workbook = new Workbook();

                // First worksheet (default)
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sales";
                sheet1.Cells["A1"].PutValue("Product");
                sheet1.Cells["B1"].PutValue("Quantity");
                sheet1.Cells["A2"].PutValue("Apple");
                sheet1.Cells["B2"].PutValue(120);

                // Second worksheet
                Worksheet sheet2 = workbook.Worksheets.Add("Inventory");
                sheet2.Cells["A1"].PutValue("Item");
                sheet2.Cells["B1"].PutValue("Stock");
                sheet2.Cells["A2"].PutValue("Banana");
                sheet2.Cells["B2"].PutValue(85);

                // Configure HTML save options.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Export each worksheet's CSS to a separate file.
                    ExportWorksheetCSSSeparately = true
                };

                // Specify a folder where the generated HTML files, CSS files and other assets will be placed.
                string outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "HtmlExport");
                saveOptions.AttachedFilesDirectory = outputFolder;

                // Ensure the output folder exists.
                Directory.CreateDirectory(outputFolder);

                // Save the workbook as HTML. The main file (output.html) will contain links to the per‑worksheet CSS files.
                string mainHtmlPath = Path.Combine(outputFolder, "output.html");
                workbook.Save(mainHtmlPath, saveOptions);

                Console.WriteLine("HTML export completed.");
                Console.WriteLine($"Main HTML file: {mainHtmlPath}");
                Console.WriteLine($"Per‑worksheet CSS files are located in: {outputFolder}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during HTML export: {ex.Message}");
            }
        }
    }

    // Entry point for the application.
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWorksheetCssSeparatelyDemo.Run();
        }
    }
}
