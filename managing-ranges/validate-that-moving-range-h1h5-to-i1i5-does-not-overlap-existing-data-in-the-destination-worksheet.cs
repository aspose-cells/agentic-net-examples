// Title: Check destination emptiness before moving range H1:H5 to I1:I5 with Aspose.Cells for .NET
// Description: C# example that creates a workbook, fills H1:H5, optionally populates I1:I5, then uses Aspose.Cells.Range.IsBlank to verify the target cells are empty before moving the source range. The program reports success or aborts to prevent overlapping data and saves the file.
// Keywords: Aspose.Cells | C# range move | IsBlank method | prevent overlapping cells | validate destination range | .NET spreadsheet manipulation | move range safely
// Common Searches: Aspose.Cells verify destination range is empty | C# move cells without overwriting data | How to check if a range is blank in Aspose.Cells | Prevent overlap when moving Excel range using Aspose | Validate range move before copying in .NET
// Developer Intent: Move H1:H5 to I1:I5 only when I1:I5 contains no data, otherwise abort the operation.
// Use Cases: Shift a column of values to the next column only if the target column is clear. | Automate data migration in spreadsheets while safeguarding existing information. | Log a warning and skip the move when the destination range already holds values.
// AI Prompts: Generate C# code that moves a range after confirming the destination is blank using Aspose.Cells. | Create a reusable method that takes source and destination range addresses and returns a boolean indicating whether the move is safe. | Provide an example that logs each overlapping cell before aborting a range move in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// C# example that creates a workbook, fills H1:H5, optionally populates I1:I5, then uses Aspose.Cells.Range.IsBlank to verify the target cells are empty before moving the source range. The program reports success or aborts to prevent overlapping data and saves the file.
class ValidateMoveRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill the source range H1:H5 with sample data
            for (int i = 0; i < 5; i++)
            {
                // Column index 7 corresponds to column H (0‑based)
                cells[i, 7].PutValue($"Src{i + 1}");
            }

            // OPTIONAL: Uncomment the following block to simulate existing data in the destination
            // for (int i = 0; i < 5; i++)
            // {
            //     // Column index 8 corresponds to column I
            //     cells[i, 8].PutValue($"Dest{i + 1}");
            // }

            // Define source and destination ranges using the Aspose.Cells.Range alias
            AsposeRange sourceRange = cells.CreateRange("H1:H5");
            AsposeRange destRange   = cells.CreateRange("I1:I5");

            // Validate that the destination range is empty (no overlapping data)
            if (destRange.IsBlank())
            {
                // Destination is clear, perform the move
                sourceRange.MoveTo(destRange.FirstRow, destRange.FirstColumn);
                Console.WriteLine("Range moved successfully.");
            }
            else
            {
                // Destination contains data, abort the move to avoid overlap
                Console.WriteLine("Destination range contains data. Move aborted to avoid overlap.");
            }

            // Save the workbook (ensure the directory is writable)
            string outputPath = "ValidateMoveRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
