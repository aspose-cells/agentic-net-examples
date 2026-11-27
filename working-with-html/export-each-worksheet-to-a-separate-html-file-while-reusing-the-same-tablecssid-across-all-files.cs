// Title: Export each worksheet of an Excel workbook to separate HTML files using a shared TableCssId with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export each worksheet to its own HTML file with the same TableCssId | How to save multiple worksheets as individual HTML pages while reusing CSS id in Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly shared TableCssId example | Generate separate HTML files for all sheets in a workbook using Aspose.Cells .NET | Set active sheet index and export to HTML with embedded Base64 images Aspose.Cells
// Tags: Aspose.Cells export worksheets to HTML | HtmlSaveOptions TableCssId shared | ExportActiveWorksheetOnly C# | embed images as Base64 Aspose.Cells | set active sheet index before HTML export

using System;
using System.IO;
using Aspose.Cells;

namespace ExportWorksheetsToHtml
{
    // The sample loads an Excel workbook, defines a common TableCssId, then loops through each worksheet, sets it as the active sheet, and uses HtmlSaveOptions (ExportActiveWorksheetOnly, ExportHiddenWorksheet, TableCssId, ExportWorksheetCSSSeparately = false, ExportImagesAsBase64 = true) to save each sheet as a self‑contained HTML file named after the sheet.
    class Program
    {
        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // 1. Determine the source workbook file.
            // -----------------------------------------------------------------
            string sourcePath;
            if (args.Length > 0 && File.Exists(args[0]))
            {
                sourcePath = args[0];
            }
            else
            {
                // Default file name – adjust as needed.
                sourcePath = "input.xlsx";
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found. Provide a valid path as a command‑line argument.");
                    return;
                }
            }

            // -----------------------------------------------------------------
            // 2. Load the workbook.
            // -----------------------------------------------------------------
            Workbook workbook;
            try
            {
                workbook = new Workbook(sourcePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // -----------------------------------------------------------------
            // 3. Prepare a common TableCssId that will be used for every HTML file.
            // -----------------------------------------------------------------
            const string sharedTableCssId = "shared-table-css";

            // -----------------------------------------------------------------
            // 4. Iterate through each worksheet and export it.
            // -----------------------------------------------------------------
            int sheetCount = workbook.Worksheets.Count;
            for (int i = 0; i < sheetCount; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Set the active sheet index so that ExportActiveWorksheetOnly works correctly.
                workbook.Worksheets.ActiveSheetIndex = i;

                // Configure HTML save options.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export only the active worksheet (the one we just set).
                    ExportActiveWorksheetOnly = true,

                    // Keep hidden worksheets (if you want to skip them, set ExportHiddenWorksheet = false).
                    ExportHiddenWorksheet = true,

                    // Use the same CSS id for all tables.
                    TableCssId = sharedTableCssId,

                    // Optional: you may want to embed CSS in each file or keep it external.
                    // Here we embed it to keep each file self‑contained.
                    ExportWorksheetCSSSeparately = false,

                    // Optional: embed images as Base64 to avoid external files.
                    ExportImagesAsBase64 = true
                };

                // Build output file name.
                // Example: Sheet1_Sales.html  (spaces in sheet name are replaced with underscores)
                string safeSheetName = sheet.Name.Replace(' ', '_');
                string outputFileName = $"Sheet{i + 1}_{safeSheetName}.html";

                try
                {
                    workbook.Save(outputFileName, htmlOptions);
                    Console.WriteLine($"Worksheet '{sheet.Name}' exported to '{outputFileName}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export worksheet '{sheet.Name}': {ex.Message}");
                }
            }

            Console.WriteLine("Export operation completed.");
        }
    }
}
