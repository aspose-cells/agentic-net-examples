// Title: How to retrieve the address of a range’s entire rows using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get full row address from a cell range | Retrieve RefersTo value after applying EntireRow property in Aspose.Cells | C# example to obtain row range address like 2:4 with Aspose.Cells | Expand a range to its entire rows and read the address using Aspose.Cells | Get entire row range reference in Aspose.Cells for .NET
// Tags: Aspose.Cells EntireRow usage | C# RefersTo property extraction | Excel full row range address | Aspose.Cells range address retrieval | C# expand range to rows

using Aspose.Cells;
using System;

// The sample creates an in‑memory workbook, defines a B2:D4 range, expands it to full rows with the EntireRow property, reads the resulting address via the RefersTo property (e.g., "2:4"), and writes the address to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (in-memory, no file needed)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a sample range (e.g., B2:D4)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("B2", "D4");

            // Obtain the rows that cover the range using the EntireRow property
            Aspose.Cells.Range entireRowRange = range.EntireRow;

            // Retrieve the address of the entire row range (e.g., "2:4")
            string address = entireRowRange.RefersTo;

            // Display the address
            Console.WriteLine("Entire row address: " + address);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
