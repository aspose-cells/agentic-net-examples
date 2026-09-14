// Title: How to freeze the first two columns (A and B) in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to freeze only columns A and B while keeping all rows scrollable. | Show the exact call to Worksheet.FreezePanes that locks the leftmost two columns and saves the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# freeze columns A and B without freezing rows | Freeze first two columns in Excel using Aspose.Cells FreezePanes method | How to lock leftmost columns in a .NET workbook with Aspose.Cells | C# example for freezing only columns in an Excel file using Aspose.Cells
// Tags: Aspose.Cells Worksheet.FreezePanes column freeze | C# freeze leftmost columns Excel workbook | Aspose.Cells freeze columns without rows | programmatic column freeze .xlsx Aspose.Cells | lock columns A B Aspose.Cells .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, accesses the first Worksheet, calls sheet.FreezePanes(0, 2, 0, 2) to freeze columns A and B while leaving rows unfrozen, and saves the file as FrozenColumns.xlsx inside a try‑catch block.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (index 0)
                Worksheet sheet = workbook.Worksheets[0];

                // Freeze the leftmost two columns (freeze columns A and B)
                // Parameters: row index, column index, total rows to freeze, total columns to freeze
                sheet.FreezePanes(0, 2, 0, 2);

                // Save the workbook to a file
                workbook.Save("FrozenColumns.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
