// Title: Count non‑empty cells in a worksheet range with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, populates sample data, defines a range (e.g., A1:C5), iterates each cell, checks that the Value is not null or empty, counts the populated cells, prints the total and optionally saves the file.
// Keywords: Aspose.Cells count non empty cells | C# Aspose.Cells range enumeration | non‑blank cell count Aspose.Cells | Aspose.Cells worksheet range | C# iterate cells Aspose.Cells
// Common Searches: Aspose.Cells count non empty cells in range | C# count filled cells Aspose.Cells | How to get number of non‑blank cells using Aspose.Cells | Count cells with data in A1:C5 Aspose.Cells | Aspose.Cells enumerate range and count
// Developer Intent: Determine how many cells contain data within a given worksheet range.
// Use Cases: Validate that a data entry block meets a minimum populated‑cell threshold before processing. | Calculate data density for reporting or conditional formatting decisions. | Skip completely empty rows or columns when exporting a selected range.
// AI Prompts: Generate a reusable C# method that receives an Aspose.Cells.Range and returns the count of non‑empty cells. | Show how to use LINQ with Aspose.Cells to count cells whose Value is neither null nor an empty string. | Explain handling of formula cells that evaluate to an empty string when counting non‑blank cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNonEmptyCountDemo
{
    // Creates a workbook, populates sample data, defines a range (e.g., A1:C5), iterates each cell, checks that the Value is not null or empty, counts the populated cells, prints the total and optionally saves the file.
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

                // Populate some sample data (some cells left empty)
                cells["A1"].PutValue("Alpha");
                cells["B2"].PutValue(123);
                cells["C3"].PutValue("Gamma");
                // Cells D4, E5 remain empty

                // Define the range to evaluate (e.g., A1:C5)
                AsposeRange range = cells.CreateRange("A1:C5");

                // Count non‑empty cells within the range
                int nonEmptyCount = 0;
                foreach (Cell cell in range)
                {
                    // A cell is considered non‑empty if its Value is not null
                    // and its string representation is not an empty string
                    if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        nonEmptyCount++;
                    }
                }

                Console.WriteLine($"Total non‑empty cells in range {range.Address}: {nonEmptyCount}");

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                workbook.Save("NonEmptyCountDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
