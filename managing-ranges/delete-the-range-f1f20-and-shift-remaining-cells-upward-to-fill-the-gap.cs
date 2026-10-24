// Title: Delete cells F1:F20 and shift remaining rows upward with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells delete range F1:F20 and shift cells up | Aspose.Cells DeleteRange example shifting rows upward in .NET | Remove vertical range F1 to F20 from Excel worksheet using Aspose.Cells | How to shift rows up after deleting a range with Aspose.Cells C# | Delete specific rows in Excel with Aspose.Cells DeleteRange ShiftType.Up
// Tags: Aspose.Cells DeleteRange ShiftType.Up | C# delete Excel range F1:F20 | remove vertical cell block Aspose.Cells | shift rows upward after deletion .NET | Excel worksheet range deletion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), accesses the first worksheet, defines the vertical range F1:F20, deletes that range while shifting the remaining cells upward, and saves the modified file as output.xlsx, with error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if it does not exist
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
                workbook.Save(inputPath);
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range F1:F20 (zero‑based indices)
            int firstRow = 0;      // Row 1
            int totalRows = 20;    // 20 rows
            int firstColumn = 5;   // Column F
            int totalColumns = 1;  // Single column

            // Delete the range and shift remaining cells upward
            sheet.Cells.DeleteRange(firstRow, totalRows, firstColumn, totalColumns, ShiftType.Up);

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
