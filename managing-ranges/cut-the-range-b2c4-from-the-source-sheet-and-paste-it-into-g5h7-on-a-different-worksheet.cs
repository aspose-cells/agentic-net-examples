// Title: Cut cells B2:C4 from a source worksheet and paste them into G5:H7 on another worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# move a range from one worksheet to another | cut and paste cells between worksheets using Aspose.Cells .NET | copy range B2:C4 to G5:H7 and clear original cells in C# | how to simulate a cut operation for an Excel range with Aspose.Cells | transfer data from Sheet1 to Sheet2 using Aspose.Cells range API
// Tags: cut range between worksheets Aspose.Cells | range copy with clear source Aspose.Cells C# | move Excel cells programmatically .NET | Aspose.Cells range transfer across sheets | clear contents after range copy Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads or creates a workbook, ensures two worksheets exist, defines the source range B2:C4 on the first sheet and the destination range G5:H7 on the second sheet, copies the source range to the destination, clears the original cells to emulate a cut operation, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook or create a new one if the source file is missing
            if (File.Exists(sourcePath))
            {
                workbook = new Workbook(sourcePath);
            }
            else
            {
                workbook = new Workbook();
                // Ensure at least two worksheets exist
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");

                // Populate sample data in the source range (B2:C4) for demonstration
                Worksheet src = workbook.Worksheets[0];
                src.Cells["B2"].PutValue("A1");
                src.Cells["C2"].PutValue("B1");
                src.Cells["B3"].PutValue("A2");
                src.Cells["C3"].PutValue("B2");
                src.Cells["B4"].PutValue("A3");
                src.Cells["C4"].PutValue("B3");

                // Save the newly created source workbook for future runs
                workbook.Save(sourcePath);
            }

            // Get source and destination worksheets
            Worksheet sourceSheet = workbook.Worksheets[0];          // e.g., "Sheet1"
            Worksheet destinationSheet = workbook.Worksheets[1];    // e.g., "Sheet2"

            // Define the source range B2:C4 on the source sheet
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange("B2", "C4");

            // Define the target range G5:H7 on the destination sheet
            AsposeRange targetRange = destinationSheet.Cells.CreateRange("G5", "H7");

            // Copy (cut) the source range into the target range
            sourceRange.Copy(targetRange);

            // Clear the original cells to complete the "cut" operation
            sourceRange.ClearContents();

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
