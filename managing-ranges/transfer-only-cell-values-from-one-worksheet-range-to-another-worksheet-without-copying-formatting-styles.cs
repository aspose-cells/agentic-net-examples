// Title: Transfer cell values only between worksheet ranges with Aspose.Cells for .NET (C#)
// Description: Shows how to copy only the raw values from a source range (A1:B3) on one worksheet to a destination range (D5:E7) on another worksheet using Aspose.Cells' CopyValue method in C#, leaving all formatting behind.
// Keywords: Aspose.Cells | CopyValue | C# | .NET | transfer cell values | range copy without formatting | worksheet range values | copy values only
// Common Searches: Aspose.Cells copy values only C# | CopyValue method example Aspose.Cells | Transfer range data without styles .NET | Copy cell values between sheets Aspose.Cells | Ignore formatting when copying cells C#
// Developer Intent: Copy only the data from a source range to a destination range while discarding all cell formatting.
// Use Cases: Build a summary sheet that pulls calculated numbers from multiple sources without altering its own layout. | Export raw dataset to a new workbook for analysis, keeping the original styling untouched. | Populate a pre‑designed template with values from a data sheet while preserving the template’s design.
// AI Prompts: Provide a C# example that uses Aspose.Cells CopyValue to move values from Sheet1!A1:B3 to Sheet2!D5:E7 without copying styles. | Explain when to use Copy, CopyStyle, and CopyValue in Aspose.Cells and the impact on formatting. | Show how to copy only cell values between worksheets in Aspose.Cells for .NET while keeping the destination formatting intact.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Shows how to copy only the raw values from a source range (A1:B3) on one worksheet to a destination range (D5:E7) on another worksheet using Aspose.Cells' CopyValue method in C#, leaving all formatting behind.
    public class TransferValuesOnly
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (source and destination will be in the same file)
                Workbook workbook = new Workbook();

                // Get the first worksheet as the source sheet and name it
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Add a second worksheet as the destination sheet and name it
                Worksheet destinationSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                destinationSheet.Name = "Destination";

                // Populate some sample data in the source range A1:B3
                Cells srcCells = sourceSheet.Cells;
                srcCells["A1"].PutValue("Item");
                srcCells["B1"].PutValue("Quantity");
                srcCells["A2"].PutValue("Apple");
                srcCells["B2"].PutValue(10);
                srcCells["A3"].PutValue("Banana");
                srcCells["B3"].PutValue(20);

                // Create Range objects for source and destination
                AsposeRange srcRange = srcCells.CreateRange("A1:B3");
                AsposeRange destRange = destinationSheet.Cells.CreateRange("D5:E7");

                // Transfer only the cell values (no formatting) from source to destination
                destRange.CopyValue(srcRange);

                // Save the workbook to a file
                string outputPath = "TransferValuesOnly.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TransferValuesOnly.Run();
        }
    }
}
