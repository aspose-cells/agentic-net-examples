// Title: Export an entire Aspose.Cells workbook to a CSV file while preserving delimiters for blank rows in C#
// AI Prompts: Generate C# code that saves a multi‑sheet Aspose.Cells workbook as a single CSV file and retains commas for empty rows. | Show how to configure TxtSaveOptions with KeepSeparatorsForBlankRow and ExportAllSheets for CSV output using Aspose.Cells.
// Common Searches: Aspose.Cells C# export workbook to CSV keep delimiters for blank rows | How to preserve empty row separators when saving Excel as CSV with Aspose.Cells .NET | TxtSaveOptions KeepSeparatorsForBlankRow example for multi‑sheet workbook | Export all worksheets to one CSV file using Aspose.Cells C#
// Tags: export workbook to CSV with Aspose.Cells | TxtSaveOptions KeepSeparatorsForBlankRow | preserve blank row delimiters CSV Aspose.Cells | ExportAllSheets CSV Aspose.Cells

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExample
{
    // The program creates a workbook, adds data with intentional blank rows, configures TxtSaveOptions (UTF‑8 encoding, comma separator, KeepSeparatorsForBlankRow=true, ExportAllSheets=true), and saves the entire workbook as a single CSV file, ensuring delimiters are kept for empty rows.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook wb = new Workbook();
            Cells cells = wb.Worksheets[0].Cells;

            // Add data with intentional blank rows to demonstrate separator preservation
            cells[0, 0].PutValue("Header1");
            cells[0, 1].PutValue("Header2");
            cells[1, 0].PutValue("Row1Col1");
            cells[1, 1].PutValue("Row1Col2");
            // Row index 2 is left blank
            cells[3, 0].PutValue("Row3Col1");
            cells[3, 1].PutValue("Row3Col2");

            // Configure TxtSaveOptions for CSV output
            TxtSaveOptions options = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF8,          // Use UTF-8 encoding
                Separator = ',',                  // Comma as the delimiter
                KeepSeparatorsForBlankRow = true, // Preserve separators for blank rows
                ExportAllSheets = true            // Export all worksheets if more than one
            };

            // Save the workbook as a CSV file with the specified options
            wb.Save("output.csv", options);
        }
    }
}
