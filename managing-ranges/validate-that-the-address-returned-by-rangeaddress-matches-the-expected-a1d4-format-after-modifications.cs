// Title: Check that Aspose.Cells Range.Address returns "A1:D4" after filling a range in C#
// Common Searches: how to compare Aspose.Cells range address with expected A1:D4 in C# | C# Aspose.Cells verify range address after populating cells | Aspose.Cells .NET validate that Range.Address equals A1:D4 | detect range address mismatch using Aspose.Cells in a workbook | exception handling for range address mismatch Aspose.Cells C#
// Tags: Aspose.Cells range address verification | C# compare Range.Address with expected string | populate range and validate address Aspose.Cells | Excel workbook range address check using Aspose.Cells | address mismatch exception Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, defines a range from A1 to D4, fills the cells with sequential numbers, retrieves the range's Address property, compares it to the expected "A1:D4" string, throws an exception if they differ, and saves the file as ModifiedRange.xlsx.
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

            // Define a range covering cells A1 to D4
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1", "D4");

            // Populate the range with sample data
            int val = 1;
            for (int r = 0; r < range.RowCount; r++)
            {
                for (int c = 0; c < range.ColumnCount; c++)
                {
                    range[r, c].PutValue(val++);
                }
            }

            // Retrieve the address of the range
            string actualAddress = range.Address;

            // Expected address string
            string expectedAddress = "A1:D4";

            // Validate that the addresses match
            if (actualAddress != expectedAddress)
            {
                throw new Exception($"Address mismatch: expected {expectedAddress}, got {actualAddress}");
            }

            // Save the workbook
            workbook.Save("ModifiedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
