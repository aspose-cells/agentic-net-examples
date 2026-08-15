// Title: Aspose.Cells .NET: Verify that Range.Address equals "A1:D4" (C#)
// Description: Creates a workbook, defines a range A1:D4 using Cells.CreateRange, reads the Range.Address property, compares it case‑insensitively with the expected "A1:D4", outputs the result, and optionally saves the file. Demonstrates how to confirm the address format before further processing.
// Keywords: Aspose.Cells C# range address | Range.Address validation | A1:D4 address check | Aspose.Cells .NET example | verify range address
// Common Searches: Aspose.Cells verify range address | Range.Address returns A1:D4 | C# Aspose.Cells check range address format | How to get range address in Aspose.Cells | Validate range address Aspose.Cells .NET
// Developer Intent: Confirm that the Range.Address property returns the exact string "A1:D4" after creating the range.
// Use Cases: Ensure correct range reference before applying formatting or formulas. | Add a sanity check in automated workbook generation pipelines. | Validate mapping between external data sources and worksheet cells.
// AI Prompts: Generate a NUnit test that asserts Range.Address is "A1:D4" for a range created with cells.CreateRange. | Show C# code to compare Range.Address case‑insensitively and log detailed differences. | Provide a method that validates the address of a dynamically sized range at runtime using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, defines a range A1:D4 using Cells.CreateRange, reads the Range.Address property, compares it case‑insensitively with the expected "A1:D4", outputs the result, and optionally saves the file. Demonstrates how to confirm the address format before further processing.
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

            // Create a range that spans from A1 to D4 (use fully qualified type to avoid ambiguity)
            Aspose.Cells.Range range = cells.CreateRange("A1:D4");

            // Retrieve the address of the created range
            string actualAddress = range.Address;

            // Expected address string
            string expectedAddress = "A1:D4";

            // Validate that the actual address matches the expected format
            if (actualAddress.Equals(expectedAddress, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Address validation passed: {actualAddress}");
            }
            else
            {
                Console.WriteLine($"Address validation failed. Expected: {expectedAddress}, Actual: {actualAddress}");
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputPath = "RangeAddressValidation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
