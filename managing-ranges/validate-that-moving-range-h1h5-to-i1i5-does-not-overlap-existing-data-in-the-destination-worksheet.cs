// Title: Validate Overlap Before Moving H1:H5 to I1:I5 with Aspose.Cells for .NET
// Description: This C# example creates a workbook, fills H1:H5, optionally places data in I3 to simulate a blocked target, defines source and destination ranges, uses the IsIntersect method to detect any overlap, moves the range only when the destination is clear, and saves the file as an XLSX document.
// Keywords: Aspose.Cells range overlap | C# IsIntersect example | move range validation Aspose.Cells | prevent data overwrite Excel | range intersect detection .NET | Aspose.Cells move range safely | Excel worksheet range move C# | Aspose.Cells range move check
// Common Searches: Aspose.Cells check if range overlaps before moving | C# move Excel range only when destination is empty | IsIntersect method usage Aspose.Cells | how to prevent overwriting cells with Aspose.Cells | validate destination range in Aspose.Cells .NET
// Developer Intent: Confirm that moving the source range H1:H5 to I1:I5 will not overwrite any existing cells in the target worksheet.
// Use Cases: Ensure a column of values can be shifted without destroying data in the adjacent column. | Safely relocate formulas or formatted cells in automated report generation. | Programmatically verify a clear target area before bulk copy‑paste operations.
// AI Prompts: Generate C# code that checks whether moving range A1:A10 to B1:B10 intersects existing data using Aspose.Cells and performs the move only if safe. | Create a reusable method that accepts source and destination addresses and returns a boolean indicating overlap via Aspose.Cells IsIntersect. | Provide an example that logs a warning instead of moving when overlap is detected, then saves the workbook.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeMoveValidation
{
    // This C# example creates a workbook, fills H1:H5, optionally places data in I3 to simulate a blocked target, defines source and destination ranges, uses the IsIntersect method to detect any overlap, moves the range only when the destination is clear, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate source range H1:H5 with sample data
                for (int i = 0; i < 5; i++)
                {
                    // Column index 7 corresponds to column H (0‑based)
                    cells[0 + i, 7].PutValue($"H{i + 1}");
                }

                // Example: populate destination cell I3 to demonstrate overlap detection
                // Comment out the following line if you want the destination to be empty
                cells[2, 8].PutValue("ExistingData"); // Row 2 = third row, Column 8 = I

                // Define source and destination ranges
                AsposeRange sourceRange = cells.CreateRange("H1:H5");
                AsposeRange destinationRange = cells.CreateRange("I1:I5");

                // Check if the source range intersects the destination range
                bool isOverlap = sourceRange.IsIntersect(destinationRange);

                if (isOverlap)
                {
                    Console.WriteLine("Cannot move the range because the destination overlaps existing data.");
                }
                else
                {
                    // Move the source range to the destination start cell (I1)
                    sourceRange.MoveTo(destinationRange.FirstRow, destinationRange.FirstColumn);
                    Console.WriteLine("Range moved successfully.");
                }

                // Save the workbook
                workbook.Save("RangeMoveValidationResult.xlsx");
                Console.WriteLine("Workbook saved as RangeMoveValidationResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
