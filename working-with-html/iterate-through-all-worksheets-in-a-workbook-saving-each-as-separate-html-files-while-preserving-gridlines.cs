// Title: Save each worksheet of an Excel workbook as a separate HTML file with gridlines using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export each worksheet to individual HTML files while keeping gridlines | How to generate separate HTML pages for every sheet in an Excel file using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions ExportGridLines true for multiple worksheets in C# | Iterate workbook worksheets and save them as HTML with safe filenames in .NET | Preserve Excel gridlines when converting worksheets to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | export worksheet to HTML with gridlines Aspose.Cells | generate separate HTML files per worksheet C# | sanitize worksheet name for file system Aspose.Cells | batch convert Excel sheets to HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, iterates through all worksheets, sets each as the active sheet, configures HtmlSaveOptions to export only the active worksheet and include gridlines, creates a file‑system‑safe HTML filename from the worksheet name, and saves each sheet as an individual HTML file while logging the process.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook from the file
                Workbook workbook = new Workbook(inputPath);

                // Iterate over a snapshot of worksheet indices to avoid collection-modification issues
                int sheetCount = workbook.Worksheets.Count;
                for (int i = 0; i < sheetCount; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];

                    // Set the current worksheet as the active one
                    workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                    // Configure HTML save options for the active sheet
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                    {
                        ExportActiveWorksheetOnly = true,
                        ExportGridLines = true
                    };

                    // Build a safe file name based on the worksheet name
                    string fileName = MakeSafeFileName(sheet.Name) + ".html";

                    // Save the active worksheet as an individual HTML file
                    workbook.Save(fileName, htmlOptions);
                    Console.WriteLine($"Saved sheet '{sheet.Name}' to '{fileName}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to replace invalid filename characters with an underscore
        static string MakeSafeFileName(string name)
        {
            foreach (char invalidChar in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(invalidChar, '_');
            }
            return name;
        }
    }
}
