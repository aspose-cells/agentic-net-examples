// Title: Validate that an Aspose.Cells range returns the A1:D4 address string in C#
// AI Prompts: Write C# code using Aspose.Cells to create a worksheet, define a range from A1 to D4, fill it with data, and assert that the Range.Address property equals "A1:D4" with a case‑insensitive check. | Show how to retrieve the A1‑style address of an Aspose.Cells range and compare it to an expected value, printing a success or failure message. | Provide a concise Aspose.Cells example that demonstrates verifying the address string of a newly created range.
// Common Searches: C# Aspose.Cells verify range address is A1:D4 | How to assert Range.Address equals expected string in Aspose.Cells | Check Excel range address format using Aspose.Cells C# example
// Tags: Aspose.Cells range address validation | C# Aspose.Range.Address comparison | A1 style address check .NET | Excel range address verification | Aspose.Cells create range A1:D4

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// // Demonstrates creating a workbook, defining a range A1:D4 with Aspose.Cells, optionally populating it, retrieving the range's Address property, and confirming it matches the expected "A1:D4" string.
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
            AsposeRange range = sheet.Cells.CreateRange("A1", "D4");

            // Fill the range with sample data (optional)
            int value = 1;
            for (int row = 0; row < range.RowCount; row++)
            {
                for (int col = 0; col < range.ColumnCount; col++)
                {
                    range[row, col].PutValue(value++);
                }
            }

            // Retrieve the address of the range in A1 style
            string address = range.Address;

            // Expected address
            string expected = "A1:D4";

            // Validate the address
            if (address.Equals(expected, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Success: Range address is '{address}'.");
            }
            else
            {
                Console.WriteLine($"Failure: Expected '{expected}' but got '{address}'.");
            }

            // Optionally save the workbook (not required for validation)
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
