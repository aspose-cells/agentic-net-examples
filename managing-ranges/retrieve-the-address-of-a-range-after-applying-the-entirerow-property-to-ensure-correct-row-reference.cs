// Title: Aspose.Cells for .NET – Retrieve Entire Row Address from a Range (C#)
// Description: Demonstrates how to create a workbook, define a range for a single cell, use the `EntireRow` property to expand the range to the full row, and obtain the row's address via the `Address` property. The example prints both the original cell address and the full‑row address, then saves the workbook.
// Keywords: Aspose.Cells EntireRow address | C# get row address from range | Aspose.Range Address property | .NET retrieve entire row | Aspose.Cells row reference | range.EntireRow C# | Aspose.Cells address string
// Common Searches: Aspose.Cells get entire row address from range | C# Aspose.Cells EntireRow property usage | How to obtain row address using Aspose.Cells | Aspose.Cells range address of whole row | Retrieve row reference from cell range .NET
// Developer Intent: Extract the address string of the full row that contains a specified range using Aspose.Cells for .NET.
// Use Cases: Log or audit the exact row location while iterating through data rows. | Apply formatting, conditional styling, or data validation to an entire row based on a single cell reference. | Generate reports that require the row address for linking or cross‑referencing within a workbook.
// AI Prompts: Generate C# code that creates a range, accesses its EntireRow property, and returns the row address using Aspose.Cells. | Explain the relationship between Aspose.Cells `Range`, `EntireRow`, and `Address` properties and show a practical example.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsEntireRowAddressDemo
{
    // Demonstrates how to create a workbook, define a range for a single cell, use the `EntireRow` property to expand the range to the full row, and obtain the row's address via the `Address` property. The example prints both the original cell address and the full‑row address, then saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Put some sample data in row 2 (index 1)
                cells["B2"].PutValue("Sample");

                // Create a range that refers to cell B2
                AsposeRange range = cells.CreateRange("B2");

                // Get the entire row that contains the range
                AsposeRange entireRow = range.EntireRow;

                // Retrieve the address of the entire row range
                string entireRowAddress = entireRow.Address;

                // Output the addresses
                Console.WriteLine("Original range address: " + range.Address);
                Console.WriteLine("Entire row address: " + entireRowAddress);

                // Save the workbook (optional, verifies that the file is created)
                string outputPath = "EntireRowAddressDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
