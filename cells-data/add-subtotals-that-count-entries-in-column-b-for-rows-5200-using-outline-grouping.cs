// Title: Count non‑empty cells in column B rows 5‑200 and write subtotal to B201 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# count non‑empty cells in column B between rows 5 and 200 | how to write a subtotal to cell B201 after counting cells with Aspose.Cells | C# iterate specific row range in Excel and save result using Aspose.Cells | save workbook as Output.xlsx after processing data with Aspose.Cells .NET | example of counting cells and inserting total in Excel using Aspose.Cells C#
// Tags: non‑empty cell counting Aspose.Cells .NET | write subtotal to specific cell Aspose.Cells | iterate row range column B Aspose.Cells C# | save workbook Output.xlsx Aspose.Cells | exception handling Aspose.Cells C# example

using System;
using System.IO;
using Aspose.Cells;

// The C# program creates a new workbook, iterates rows 5‑200 in column B counting cells that contain a non‑empty string, writes the total count to cell B201, and saves the file as Output.xlsx, with basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for which we want to count non‑empty entries in column B (zero‑based index 1)
            int startRow = 5;   // zero‑based index (row 6 in Excel)
            int endRow = 200;   // zero‑based index (row 201 in Excel)
            int columnIndex = 1; // column B

            // Count non‑empty cells in the specified range
            int count = 0;
            for (int row = startRow; row <= endRow; row++)
            {
                var cell = sheet.Cells[row, columnIndex];
                if (cell != null && !string.IsNullOrEmpty(cell.StringValue))
                {
                    count++;
                }
            }

            // Write the count result below the range (e.g., cell B201)
            int resultRow = endRow + 1;
            sheet.Cells[resultRow, columnIndex].PutValue(count);

            // Prepare output file path
            string outputPath = "Output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
