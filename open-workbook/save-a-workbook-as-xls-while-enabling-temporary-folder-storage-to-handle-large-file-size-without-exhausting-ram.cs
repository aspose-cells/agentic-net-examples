// Title: Save a large Excel workbook to XLS format with disk‑based memory setting using Aspose.Cells in C#
// Common Searches: how to export a huge workbook to .xls with Aspose.Cells in C# without running out of RAM | Aspose.Cells C# memory preference option for saving large Excel 97‑2003 files | using temporary folder storage in Aspose.Cells to write big Excel files to XLS | C# example for saving 50,000 rows to an XLS file with disk‑based memory setting in Aspose.Cells | reduce memory consumption when saving large workbooks to Excel97To2003 format with Aspose.Cells
// Tags: disk‑based storage Aspose.Cells C# | memory setting Aspose.Cells | export workbook to Excel97To2003 Aspose.Cells | temporary folder usage Aspose.Cells large workbook | low‑memory XLS save Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a Workbook, switches to disk‑based storage by setting MemorySetting to MemoryPreference, optionally fills 50,000 rows to simulate a large file, and saves the result as an Excel 97‑2003 (.xls) document using SaveFormat.Excel97To2003, all wrapped in a try‑catch block to handle errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Use disk-based storage for large files to reduce RAM consumption
            workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

            // (Optional) Populate the workbook with data to simulate a large file
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 50000; row++)
            {
                sheet.Cells[row, 0].PutValue($"Row {row + 1}");
                sheet.Cells[row, 1].PutValue(row);
            }

            // Save the workbook as an Excel 97-2003 file
            workbook.Save("LargeFileOutput.xls", SaveFormat.Excel97To2003);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
