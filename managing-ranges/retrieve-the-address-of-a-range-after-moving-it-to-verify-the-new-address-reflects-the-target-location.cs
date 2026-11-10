// Title: How to retrieve the new address of a moved range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get range address after using MoveTo | C# Aspose.Cells retrieve RefersTo value of a moved range | How to verify new range location after moving cells with Aspose.Cells | Aspose.Cells MoveTo method returns updated range reference | Get updated range address programmatically in Aspose.Cells .NET
// Tags: range.MoveTo usage Aspose.Cells | RefersTo property for moved range | updated range address verification | Aspose.Cells range relocation example | C# workbook save after range move

using System;
using Aspose.Cells;

// The example creates a workbook, fills cells A1:B2, defines a range, moves it to C3 using the MoveTo method, reads the new address via the RefersTo property, prints it, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data in the original range A1:B2
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);

            // Create a range that covers A1:B2 (use fully qualified type to avoid ambiguity)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1", "B2");

            // Move the range to the target location C3 (zero‑based indices: row 2, column 2)
            range.MoveTo(2, 2);

            // Retrieve the address of the moved range
            string movedAddress = range.RefersTo; // Expected format: "C3:D4"
            Console.WriteLine("New address: " + movedAddress);

            // Save the workbook
            string outputPath = "MovedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
