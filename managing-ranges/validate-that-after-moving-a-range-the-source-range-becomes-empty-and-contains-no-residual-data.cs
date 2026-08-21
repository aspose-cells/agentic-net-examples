// Title: Aspose.Cells .NET – Verify Source Range Is Empty After Range.MoveTo
// Description: C# example that creates a workbook, fills A1:B2, moves the range to A3:B4 with Range.MoveTo, then uses Range.IsBlank to confirm the original cells are cleared and the destination holds the data before saving the file.
// Keywords: Aspose.Cells | C# | .NET | Range.MoveTo | Range.IsBlank | validate moved range | source range empty | worksheet range verification | Aspose.Cells example
// Common Searches: Aspose.Cells check if source range is blank after MoveTo | C# verify range move clears original cells | How to test Range.MoveTo in Aspose.Cells | Aspose.Cells .NET validate range relocation | Range.IsBlank after moving cells
// Developer Intent: Confirm that calling Range.MoveTo removes all data from the original range.
// Use Cases: Automated unit tests that assert data integrity after moving a range within a worksheet. | Debugging scripts to ensure no residual values remain in the source area after a range relocation. | Generating reports where ranges are repositioned and the original cells must be cleared before saving.
// AI Prompts: Generate an xUnit test that moves a range with Aspose.Cells and asserts the source range is blank using Range.IsBlank. | Write a reusable C# method that moves any range to a new address and returns true if the original cells are empty. | Explain how Range.MoveTo handles merged cells and how to verify that merged source cells are cleared after the move.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeMoveValidation
{
    // C# example that creates a workbook, fills A1:B2, moves the range to A3:B4 with Range.MoveTo, then uses Range.IsBlank to confirm the original cells are cleared and the destination holds the data before saving the file.
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

                // Create a source range A1:B2 and put sample data into it
                AsposeRange sourceRange = cells.CreateRange("A1:B2");
                sourceRange[0, 0].PutValue("A1");
                sourceRange[0, 1].PutValue("B1");
                sourceRange[1, 0].PutValue("A2");
                sourceRange[1, 1].PutValue("B2");

                // Move the source range down by two rows to A3:B4
                // MoveTo expects zero‑based row and column indices
                sourceRange.MoveTo(sourceRange.FirstRow + 2, sourceRange.FirstColumn);

                // Verify that the original location (A1:B2) is now empty
                AsposeRange originalLocation = cells.CreateRange("A1:B2");
                bool isBlank = originalLocation.IsBlank();
                Console.WriteLine($"Original range A1:B2 is blank after move: {isBlank}");

                // Verify that the destination range contains the moved data
                AsposeRange destinationRange = cells.CreateRange("A3:B4");
                bool destHasData = !destinationRange.IsBlank();
                Console.WriteLine($"Destination range A3:B4 has data after move: {destHasData}");

                // Save the workbook
                string outputPath = "RangeMoveValidation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
