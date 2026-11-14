// Title: How to duplicate a header row and freeze the top rows using Cells.CopyRow and FreezePanes in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy first row to second row and freeze panes | How to use Cells.CopyRow to duplicate a header then freeze rows in Excel with Aspose.Cells | C# code example for copying a header row and applying FreezePanes in Aspose.Cells | Freeze top rows after copying header row using Aspose.Cells .NET API | Duplicate Excel header and lock it while scrolling with Aspose.Cells in C#
// Tags: duplicate header row Aspose.Cells C# | freeze top rows Aspose.Cells .NET | Cells.CopyRow method usage | FreezePanes API example | copy row and freeze panes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing workbook, copies the first worksheet's header row (row 0) to row 1 using Cells.CopyRow, freezes the top two rows with the FreezePanes method, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Copy header row (row 0) to row 1
                // Use the overload that requires the source Cells object
                sheet.Cells.CopyRow(sheet.Cells, 0, 1);

                // Freeze the top two rows (rows 0 and 1)
                // Worksheet provides a FreezePanes method directly
                sheet.FreezePanes(2, 0, 2, 0);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
