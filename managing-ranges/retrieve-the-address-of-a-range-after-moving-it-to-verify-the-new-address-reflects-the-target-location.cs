// Title: How to get the updated address of a moved range with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# retrieve new range address after using MoveTo | Get updated RefersTo value after moving a range with Aspose.Cells | C# example to verify moved range address in an Excel workbook using Aspose.Cells | How to obtain the new location of a range after MoveTo in Aspose.Cells .NET | Aspose.Cells range relocation address verification code sample
// Tags: Aspose.Cells MoveTo range address | C# Aspose.Cells RefersTo property | Excel range relocation with Aspose.Cells | .NET Aspose.Cells range address verification | Aspose.Cells moved range address retrieval

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, fills cells A1:C3, defines a range, moves it to E5 using the MoveTo method, and then reads the RefersTo property to obtain the new address (e.g., "E5:G7").
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a sample range A1:C3 with data
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(3);
            sheet.Cells["A2"].PutValue(4);
            sheet.Cells["B2"].PutValue(5);
            sheet.Cells["C2"].PutValue(6);
            sheet.Cells["A3"].PutValue(7);
            sheet.Cells["B3"].PutValue(8);
            sheet.Cells["C3"].PutValue(9);

            // Define the source range (A1:C3)
            AsposeRange sourceRange = sheet.Cells.CreateRange("A1:C3");

            // Move the range to a new location, e.g., starting at E5 (row 4, column 4 in zero‑based indices)
            sourceRange.MoveTo(4, 4);

            // Retrieve the new address of the moved range
            string newAddress = sourceRange.RefersTo; // e.g., "E5:G7"

            // Output the new address to verify the move
            Console.WriteLine("New address of the moved range: " + newAddress);

            // Save the workbook to inspect the result
            workbook.Save("MovedRangeOutput.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
