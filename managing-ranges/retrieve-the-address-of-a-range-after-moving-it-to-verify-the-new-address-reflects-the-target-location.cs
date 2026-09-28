// Title: Move a cell range to a new location and read its updated RefersTo address using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells Range.MoveTo to relocate the range A1:B2 to C3 and then output the Range.RefersTo string in C#. | After moving a range with Range.MoveTo, retrieve the RefersTo property to confirm the new address in a .NET workbook.
// Common Searches: Aspose.Cells C# get address of a range after moving it | How to retrieve new range reference after Range.MoveTo in .NET | C# Aspose.Cells move range to another cell and verify RefersTo | Updated range address after relocating cells with Aspose.Cells
// Tags: Range.MoveTo method Aspose.Cells | RefersTo property after moving range | relocating cell range .NET | updated range address C# Aspose.Cells | move range and verify address

using Aspose.Cells;
using System;

// // Demonstrates creating a workbook, filling A1:B2, moving that range to C3 with Range.MoveTo, and reading the RefersTo property to display the range's new address.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill some data in the source range A1:B2
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["B2"].PutValue(4);

            // Define the source range (use fully qualified Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange("A1:B2");

            // Move the range to target location C3 (row index 2, column index 2 – zero‑based)
            sourceRange.MoveTo(2, 2);

            // Retrieve the new address of the moved range
            string newAddress = sourceRange.RefersTo;

            // Output the new address
            Console.WriteLine("New address of the moved range: " + newAddress);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
