// Title: How to retrieve the row range address (e.g., "2:4") from a cell range using the EntireRow property in Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a worksheet, defines a cell range, applies the EntireRow property, and outputs the RefersTo address of the resulting row range with Aspose.Cells. | Show how to extract the start and end row numbers of a specific range (e.g., B2:D4) by using the EntireRow property in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# get row numbers of a range using EntireRow | How to obtain the address of rows covering a cell range in Aspose.Cells | Retrieve RefersTo string for entire rows from a selected range in .NET | Example of using Range.EntireRow to find row range address in Aspose.Cells
// Tags: Aspose.Cells range entirerow address | C# retrieve row range RefersTo Aspose.Cells | convert cell range to full rows Aspose.Cells | Aspose.Cells get covering rows from range | row address extraction using EntireRow property

using Aspose.Cells;
using System;

// Creates a workbook, defines a range B2:D4, uses the EntireRow property to get the rows that span the range, and prints the row address (e.g., "2:4") via the RefersTo property.
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

            // Define a range (for example, cells B2:D4)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("B2", "D4");

            // Apply the EntireRow property to get the rows that cover the range
            Aspose.Cells.Range entireRowRange = range.EntireRow;

            // Retrieve the address of the entire row range (e.g., "2:4")
            string address = entireRowRange.RefersTo;

            // Output the address
            Console.WriteLine("Entire row address: " + address);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
