// Title: Shift a cell range five rows upward and copy its formatting back to the original range using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that creates a range offset five rows above an existing range and copies the offset range’s style to the original cells. | Show how to clamp the offset start row to zero and transfer formatting from the offset range to the source range in a .NET Excel workbook.
// Common Searches: Aspose.Cells offset range upward by 5 rows and copy formatting | C# copy style from one Excel range to another using Aspose.Cells | How to shift a range up and apply its formatting in .NET Excel library | Create range with same size at different row index Aspose.Cells C#
// Tags: offset range rows Aspose.Cells | copy range style Aspose.Cells | shift range upward .NET Excel | clone formatting between ranges C# | create range same dimensions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program loads (or creates) an Excel workbook, defines a source range (e.g., A1:C10), creates an offset range five rows above (clamped to the first row), copies the formatting from the offset range back to the original range, and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a blank workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var newWorkbook = new Workbook();
                newWorkbook.Save(inputPath);
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the original range (example: A1:C10).
            AsposeRange originalRange = worksheet.Cells.CreateRange("A1:C10");

            // Calculate the starting row for the offset range (5 rows upward).
            int offsetStartRow = originalRange.FirstRow - 5;
            if (offsetStartRow < 0) offsetStartRow = 0;

            // Create the offset range with the same size as the original range.
            AsposeRange offsetRange = worksheet.Cells.CreateRange(
                offsetStartRow,
                originalRange.FirstColumn,
                originalRange.RowCount,
                originalRange.ColumnCount);

            // Copy formatting (style) from the offset range back to the original range.
            originalRange.CopyStyle(offsetRange);

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
