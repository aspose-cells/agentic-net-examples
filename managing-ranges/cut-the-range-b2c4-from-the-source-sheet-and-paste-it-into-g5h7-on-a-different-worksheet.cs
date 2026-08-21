// Title: C# Aspose.Cells: Cut range B2:C4 and paste to G5:H7 on another worksheet
// Description: Demonstrates how to create a workbook, fill cells B2:C4 on a source sheet, add a destination sheet, and use InsertCutCells with ShiftType.Down to cut the range and paste it starting at cell G5 (G5:H7) on the destination sheet, then save the file.
// Keywords: Aspose.Cells | C# | .NET | cut range | paste range | InsertCutCells | ShiftType.Down | move cells between worksheets | cut and paste cells | range B2:C4 | cell G5
// Common Searches: Aspose.Cells cut range and paste to another sheet | InsertCutCells C# example | How to move cells B2:C4 to G5:H7 with Aspose.Cells | Cut and paste cells across worksheets .NET | ShiftType.Down effect in Aspose.Cells
// Developer Intent: Cut the range B2:C4 from a source worksheet and paste it into G5:H7 on a different worksheet using Aspose.Cells for .NET.
// Use Cases: Transfer a data block from a raw data sheet to a formatted report sheet without losing layout. | Reposition a calculated table onto a summary sheet while automatically shifting existing cells down. | Programmatically reorganize workbook content by cutting and pasting ranges across multiple worksheets.
// AI Prompts: Generate C# code that uses Aspose.Cells to cut the range B2:C4 from one worksheet and paste it into G5:H7 on another worksheet. | Explain the InsertCutCells method parameters, including row/column indexes and the impact of different ShiftType values. | Show how to handle errors when cutting and pasting ranges with Aspose.Cells in a .NET application.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCutPasteDemo
{
    // Demonstrates how to create a workbook, fill cells B2:C4 on a source sheet, add a destination sheet, and use InsertCutCells with ShiftType.Down to cut the range and paste it starting at cell G5 (G5:H7) on the destination sheet, then save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();

                // Access the first worksheet as the source sheet
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Populate the source range B2:C4 with sample data
                sourceSheet.Cells["B2"].PutValue("B2");
                sourceSheet.Cells["C2"].PutValue("C2");
                sourceSheet.Cells["B3"].PutValue("B3");
                sourceSheet.Cells["C3"].PutValue("C3");
                sourceSheet.Cells["B4"].PutValue("B4");
                sourceSheet.Cells["C4"].PutValue("C4");

                // Create a second worksheet as the destination sheet
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // Define the range to cut from the source sheet (B2:C4)
                AsposeRange cutRange = sourceSheet.Cells.CreateRange("B2:C4");

                // Insert the cut range into the destination sheet at G5 (row index 4, column index 6)
                // ShiftType.Down is used to shift existing cells down if needed
                destSheet.Cells.InsertCutCells(cutRange, 4, 6, ShiftType.Down);

                // Save the workbook (lifecycle save)
                workbook.Save("CutPasteResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
