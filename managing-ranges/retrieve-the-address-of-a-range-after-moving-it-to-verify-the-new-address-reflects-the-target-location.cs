// Title: Get Updated Range Address After MoveTo with Aspose.Cells for .NET
// Description: Creates a workbook, defines a range (A1:B2), moves it one row down using Range.MoveTo, and reads the new Address property to verify the new location, with optional saving of the file.
// Keywords: Aspose.Cells | C# range MoveTo | range address after move | Range.MoveTo example | Aspose.Cells Get Address | Excel range relocation .NET | Aspose.Cells Address property | C# Aspose.Cells range manipulation
// Common Searches: Aspose.Cells get range address after moving | C# MoveTo range new address | Aspose.Cells Range.MoveTo returns new location | retrieve updated range address Aspose.Cells | range.Address after MoveTo .NET
// Developer Intent: Obtain the new address of a cell range after it has been moved with Aspose.Cells.
// Use Cases: Validate that a MoveTo operation repositioned data by comparing original and new range addresses. | Use the updated Address value to build formulas or references that depend on the moved range. | Log range addresses during automated workbook transformations for debugging and audit trails.
// AI Prompts: Write C# code that moves a range with Aspose.Cells and prints the updated range address. | Explain how the Range.Address property reflects the new location after calling MoveTo. | Show how to compare original and new range addresses to confirm a successful MoveTo operation.

using System;
using Aspose.Cells;

namespace AsposeCellsRangeMoveDemo
{
    // Creates a workbook, defines a range (A1:B2), moves it one row down using Range.MoveTo, and reads the new Address property to verify the new location, with optional saving of the file.
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

                // Create a sample range (A1:B2) and put some data
                Aspose.Cells.Range range = cells.CreateRange("A1", "B2");
                range[0, 0].PutValue("A1");
                range[0, 1].PutValue("B1");
                range[1, 0].PutValue("A2");
                range[1, 1].PutValue("B2");

                // Display original address
                Console.WriteLine("Original range address: " + range.Address);

                // Move the range down by one row (to A2:B3)
                range.MoveTo(range.FirstRow + 1, range.FirstColumn);

                // Retrieve and display the new address after moving
                string newAddress = range.Address;
                Console.WriteLine("New range address after MoveTo: " + newAddress);

                // Save the workbook (optional, just to verify the move visually)
                string outputPath = "MovedRangeDemo.xlsx";
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
