// Title: How to verify that a source range is cleared after using MoveTo in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# check if original cells are empty after moving a range | how to validate source range cleared when using MoveTo in Aspose.Cells | C# code to ensure source range becomes null after range relocation with Aspose.Cells | verify that MoveTo empties the source cells in Aspose.Cells .NET | example of testing source range emptiness after moving a range in Aspose.Cells
// Tags: Aspose.Cells MoveTo source range clearance | C# range relocation empty cells verification | cell IsNull detection after range move Aspose.Cells | Aspose.Cells workbook save after range transfer | range move validation example Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, fills A1:C3 with numbers, defines source and destination ranges, moves the source range to E5:G7 using MoveTo, then iterates over the original source range to confirm each cell's type is IsNull, outputs the validation result, and saves the workbook as MovedRange.xlsx.
class MoveRangeValidation
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source range A1:C3 with sample data
            int value = 1;
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue(value++);
                }
            }

            // Define source and destination ranges (use fully qualified Aspose.Cells.Range)
            Aspose.Cells.Range sourceRange = cells.CreateRange("A1:C3");
            Aspose.Cells.Range destRange = cells.CreateRange("E5:G7");

            // Move the source range to the destination range (using row/column indices)
            try
            {
                sourceRange.MoveTo(destRange.FirstRow, destRange.FirstColumn);
            }
            catch (Exception moveEx)
            {
                Console.WriteLine($"Error during MoveTo: {moveEx.Message}");
                return;
            }

            // Validate that the original source cells are now empty
            bool isSourceEmpty = true;
            foreach (Cell cell in sourceRange)
            {
                if (cell.Type != CellValueType.IsNull)
                {
                    isSourceEmpty = false;
                    break;
                }
            }

            // Output validation result
            Console.WriteLine(isSourceEmpty
                ? "Source range is empty after move."
                : "Source range still contains data after move.");

            // Save the workbook to verify visually
            string outputPath = "MovedRange.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
