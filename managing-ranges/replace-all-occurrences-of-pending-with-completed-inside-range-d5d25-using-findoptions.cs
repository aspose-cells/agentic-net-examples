// Title: Replace "Pending" with "Completed" in cells D5:D25 using Aspose.Cells FindOptions (C#)
// Common Searches: C# code to replace text in a specific Excel range with Aspose.Cells FindOptions | Aspose.Cells find and replace values in range D5:D25 example | How to update cells containing 'Pending' to 'Completed' using Aspose.Cells in C# | Search for exact cell value and replace it in Excel with Aspose.Cells C# | Programmatically replace column D values in an Excel workbook using Aspose.Cells
// Tags: Aspose.Cells FindOptions replace cell values | C# Excel range text replacement Aspose.Cells | search and replace specific range D5:D25 Aspose.Cells | exact match cell value update Aspose.Cells | programmatic Excel column D value modification C#

using System;
using System.IO;
using Aspose.Cells;

// This C# example loads an existing workbook, defines the range D5:D25 on the first worksheet, and uses Aspose.Cells FindOptions to locate every cell whose value exactly matches "Pending". Each matching cell is updated to "Completed" and the workbook is saved as output.xlsx. The code includes input file validation and exception handling.
class ReplacePendingWithCompleted
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the search range D5:D25 (zero‑based indices)
            int startRow = 4;   // Row 5
            int endRow = 24;    // Row 25
            int columnIndex = 3; // Column D

            // Configure FindOptions to search within cell values (exact match)
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values
                // Default LookAtType is Whole, which matches the entire cell content
            };

            // Start searching from the first cell in the range
            Cell startCell = worksheet.Cells[startRow, columnIndex];
            Cell foundCell = worksheet.Cells.Find("Pending", startCell, findOptions);

            // Iterate through all matches within the defined range
            while (foundCell != null &&
                   foundCell.Row >= startRow && foundCell.Row <= endRow &&
                   foundCell.Column == columnIndex)
            {
                foundCell.PutValue("Completed");

                // Move start cell to the next row to continue searching
                int nextRow = foundCell.Row + 1;
                if (nextRow > endRow)
                    break; // No more rows in the range

                startCell = worksheet.Cells[nextRow, columnIndex];
                foundCell = worksheet.Cells.Find("Pending", startCell, findOptions);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
