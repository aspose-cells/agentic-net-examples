// Title: Freeze the first two columns of an Excel worksheet using Aspose.Cells in C#
// Common Searches: c# aspose.cells freeze first two columns without freezing rows | how to lock columns A and B in an Excel file using Aspose.Cells FreezePanes | example of keeping identifier columns static while scrolling horizontally with Aspose.Cells | aspose.cells freeze panes to prevent horizontal scroll from moving first columns | c# code to freeze columns only in a workbook using Aspose.Cells
// Tags: Aspose.Cells FreezePanes column freeze | C# Excel column freeze with Aspose | static identifier columns Aspose.Cells | horizontal scrolling column lock Excel | freeze columns A B Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing workbook, accesses the first worksheet, and calls FreezePanes(0, 2, 0, 2) to freeze columns A and B while leaving rows unfrozen, then saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Freeze the first two columns (A and B). No rows are frozen.
                // Parameters: row, column, totalRows, totalColumns
                sheet.FreezePanes(0, 2, 0, 2);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
