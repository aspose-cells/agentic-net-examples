// Title: Export Worksheets to Separate HTML Files with Preserved Links using Aspose.Cells (C#)
// Description: Demonstrates how to save each worksheet of an Aspose.Cells workbook as an individual HTML file while keeping inter‑worksheet hyperlinks functional. The example implements a custom IFilePathProvider that returns "<SheetName>.html", configures HtmlSaveOptions with ExportActiveWorksheetOnly, and writes the files to disk.
// Keywords: Aspose.Cells HTML export | C# export worksheets to HTML | IFilePathProvider example | preserve worksheet hyperlinks | ExportActiveWorksheetOnly | separate HTML files per sheet | Aspose.Cells custom file naming
// Common Searches: Aspose.Cells export each worksheet to separate HTML file | keep hyperlinks between worksheets when saving as HTML | custom IFilePathProvider for HTML export Aspose.Cells | C# save workbook as multiple HTML pages | how to use ExportActiveWorksheetOnly Aspose.Cells
// Developer Intent: Create individual HTML pages for every worksheet and ensure that hyperlinks between sheets continue to work after export.
// Use Cases: Publish a multi‑sheet Excel report as a web‑ready set of pages with navigation links. | Generate per‑sheet documentation for a portal where each section is a separate HTML file. | Automate batch conversion of workbooks to HTML with a naming scheme controlled by a custom provider.
// AI Prompts: Show how to modify CustomFilePathProvider to store HTML files in a subfolder while preserving links. | Provide C# code that exports only selected worksheets to separate HTML files with custom filenames. | Explain how to add a hyperlink that points to a specific cell in another worksheet after exporting to separate HTML files.

using System;
using System.IO;
using Aspose.Cells;

namespace ExportWorksheetsToSeparateHtml
{
    // Custom implementation of IFilePathProvider.
    // Returns a file name for each worksheet so that links between worksheets are preserved.
    // Demonstrates how to save each worksheet of an Aspose.Cells workbook as an individual HTML file while keeping inter‑worksheet hyperlinks functional. The example implements a custom IFilePathProvider that returns "<SheetName>.html", configures HtmlSaveOptions with ExportActiveWorksheetOnly, and writes the files to disk.
    internal class CustomFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Each worksheet will be saved as "<SheetName>.html" in the same directory as the main file.
            return $"{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with three worksheets.
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Populate some data.
                workbook.Worksheets["Sheet1"].Cells["A1"].PutValue("Data in Sheet1");
                workbook.Worksheets["Sheet2"].Cells["A1"].PutValue("Data in Sheet2");
                workbook.Worksheets["Sheet3"].Cells["A1"].PutValue("Data in Sheet3");

                // Add a hyperlink in Sheet1 that points to Sheet2.
                // The hyperlink will be updated automatically to refer to the correct HTML file.
                Worksheet sheet1 = workbook.Worksheets["Sheet1"];
                try
                {
                    // Add hyperlink at cell C3 (row index 2, column index 2 – zero based indexing).
                    // totalRows = 1, totalColumns = 1 for a single cell.
                    // Use overload with 5 parameters (screen tip set separately if needed).
                    sheet1.Hyperlinks.Add(2, 2, 1, 1, "Sheet2!A1");
                    // Optionally set display text and screen tip.
                    sheet1.Cells["C3"].PutValue("Go to Sheet2");
                    // Set screen tip if desired.
                    if (sheet1.Hyperlinks.Count > 0)
                    {
                        sheet1.Hyperlinks[0].ScreenTip = "Go to Sheet2";
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add hyperlink: {ex.Message}");
                }

                // Configure HTML save options.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Export each worksheet separately.
                    ExportActiveWorksheetOnly = true,
                    // Use the custom provider to generate file names.
                    FilePathProvider = new CustomFilePathProvider()
                };

                // Determine output file path and ensure directory exists.
                string outputFile = "Workbook.html";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile)) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook. The main file name is arbitrary; separate files will be created per worksheet.
                workbook.Save(outputFile, saveOptions);

                Console.WriteLine("Worksheets exported to separate HTML files with preserved links.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
