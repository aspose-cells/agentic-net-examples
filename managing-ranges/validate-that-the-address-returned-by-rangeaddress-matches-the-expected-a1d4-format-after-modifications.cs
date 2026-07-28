// Title: Validate Aspose.Cells Range.Address returns "A1:D4" after creating a range (C#)
// Description: Creates a workbook, defines a range A1:D4 with Cells.CreateRange, reads the Range.Address property, compares it to the expected "A1:D4" string (case‑insensitive), outputs the result, and saves the file.
// Keywords: Aspose.Cells | Range.Address | C# | A1:D4 | CreateRange | range address validation | unit test | address comparison
// Common Searches: Aspose.Cells verify range address C# | Range.Address returns wrong format | how to compare Aspose range address with expected value | validate range address after creating range in Aspose.Cells | C# Aspose.Cells check A1:D4 address
// Developer Intent: Confirm that the Range.Address property exactly matches the expected "A1:D4" address for a newly created range.
// Use Cases: Unit‑test verification of range addresses before applying formatting or formulas. | Automated workbook generation that requires precise range references. | Debugging dynamic range adjustments where address mismatches can cause errors.
// AI Prompts: Generate an xUnit test that asserts Range.Address equals "A1:D4" after calling Cells.CreateRange. | Provide a logging snippet that records the expected and actual range addresses when they differ. | Show how to expand the range to "E5" and re‑validate that Range.Address reflects the new address.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeAddressValidation
{
    // Creates a workbook, defines a range A1:D4 with Cells.CreateRange, reads the Range.Address property, compares it to the expected "A1:D4" string (case‑insensitive), outputs the result, and saves the file.
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

                // Create a range that covers cells A1 to D4
                AsposeRange range = cells.CreateRange("A1:D4");

                // Retrieve the address of the created range
                string actualAddress = range.Address;

                // Expected address string
                const string expectedAddress = "A1:D4";

                // Validate that the actual address matches the expected address
                if (actualAddress.Equals(expectedAddress, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Range address validation succeeded: " + actualAddress);
                }
                else
                {
                    Console.WriteLine($"Range address validation failed. Expected: {expectedAddress}, Actual: {actualAddress}");
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save("RangeAddressValidation.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
