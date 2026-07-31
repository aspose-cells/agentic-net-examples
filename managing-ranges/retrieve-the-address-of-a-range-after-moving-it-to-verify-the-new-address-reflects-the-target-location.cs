// Title: Get Updated Range Address After MoveTo with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, define a range, move it with Range.MoveTo, and read the Range.Address property to obtain the new cell address, confirming the range object updates automatically.
// Keywords: Aspose.Cells | C# | .NET | Range.MoveTo | range address | updated range address | move range | cell range relocation | Range.Address | Aspose.Cells example | retrieve range address
// Common Searches: Aspose.Cells get range address after MoveTo | C# move range and read new address | How to retrieve new address of moved range Aspose.Cells | Range.MoveTo updates address | Aspose.Cells .NET range relocation address
// Developer Intent: Find the new address of a range after it has been moved with Aspose.Cells.
// Use Cases: Confirm that a programmatically shifted range lands on the intended cells before exporting. | Log the new range location for debugging or audit trails. | Refresh formulas, charts, or data validations that depend on the moved range.
// AI Prompts: Write C# code that moves a range by N rows and columns using Aspose.Cells and prints the updated address each time. | Explain why Range.Address reflects the new position after calling MoveTo in Aspose.Cells. | Show how to capture and store the new range address after moving a range for later use.

using System;
using Aspose.Cells;

namespace AsposeCellsRangeMoveDemo
{
    // Shows how to create a workbook, define a range, move it with Range.MoveTo, and read the Range.Address property to obtain the new cell address, confirming the range object updates automatically.
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

                // Create a sample range A1:B2 and put some data
                Aspose.Cells.Range originalRange = cells.CreateRange("A1", "B2");
                originalRange[0, 0].PutValue("A1");
                originalRange[0, 1].PutValue("B1");
                originalRange[1, 0].PutValue("A2");
                originalRange[1, 1].PutValue("B2");

                // Display the original address
                Console.WriteLine("Original range address: " + originalRange.Address);

                // Move the range down by 1 row (to A2:B3)
                originalRange.MoveTo(originalRange.FirstRow + 1, originalRange.FirstColumn);

                // After moving, the same Range object reflects the new location
                Console.WriteLine("New range address after MoveTo: " + originalRange.Address);

                // Save the workbook (optional, just to verify the move visually)
                workbook.Save("RangeMoveResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
