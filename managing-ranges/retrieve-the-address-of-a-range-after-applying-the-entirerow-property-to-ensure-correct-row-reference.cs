// Title: Get the Entire Row Address of a Range with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, define a range starting at a specific cell, use the EntireRow property to expand the range to the full row, retrieve the row's address string, and optionally save the file.
// Keywords: Aspose.Cells EntireRow address | C# get whole row range address | Aspose.Range EntireRow property | retrieve row address from cell range | .NET Aspose.Cells range address | Aspose.Cells get row address | EntireRow property example | Aspose.Cells C# range manipulation
// Common Searches: Aspose.Cells get entire row address from range | C# EntireRow property returns row address | How to obtain row address after creating a range in Aspose.Cells | Aspose.Cells retrieve whole row range address | Get address of entire row using Aspose.Cells for .NET
// Developer Intent: Obtain the address string of the full row that corresponds to a given cell range.
// Use Cases: Log original and full‑row addresses for debugging or audit trails. | Apply formatting, formulas, or data validation to an entire row based on its address. | Pass the row address to downstream processes such as reporting or automated workflows.
// AI Prompts: Write C# code with Aspose.Cells that creates a range at a specified cell, accesses the EntireRow property, and prints the row's address. | Explain how the EntireRow property works in Aspose.Cells and how to retrieve its Address value in .NET. | Provide an example that creates a range at B5, gets the entire row range, and uses the address to apply a style to that row.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a workbook, define a range starting at a specific cell, use the EntireRow property to expand the range to the full row, retrieve the row's address string, and optionally save the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a range that starts at cell A2
            AsposeRange range = cells.CreateRange("A2");

            // Display the address of the original range
            Console.WriteLine("Original range address: " + range.Address);

            // Use the EntireRow property to obtain a range that represents the whole row
            AsposeRange entireRow = range.EntireRow;

            // Retrieve and display the address of the entire row range
            Console.WriteLine("Entire row address: " + entireRow.Address);

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputPath = "EntireRowAddressDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
