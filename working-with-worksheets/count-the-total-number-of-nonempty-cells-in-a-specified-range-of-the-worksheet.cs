// Title: Count non‑empty cells in a worksheet range with Aspose.Cells for .NET (C#)
// Description: A concise C# example that creates a workbook, fills selected cells, defines a range (e.g., "A1:C5"), iterates the Aspose.Cells.Range collection, checks each Cell.Value for null or empty string, increments a counter for populated cells, writes the total to the console, and optionally saves the file.
// Keywords: Aspose.Cells count non empty cells | C# count cells in range Aspose.Cells | Aspose.Cells iterate range | check empty cell Aspose.Cells | worksheet cell count Aspose.Cells | Aspose.Cells .NET range enumeration
// Common Searches: count non empty cells Aspose.Cells C# | how to count filled cells in Excel range using Aspose.Cells | Aspose.Cells iterate over range and count values | determine number of populated cells in a worksheet range Aspose.Cells | C# Aspose.Cells count cells ignoring blanks
// Developer Intent: The developer wants to know how many cells in a specified worksheet range contain data.
// Use Cases: Validate that a data entry block contains the expected number of filled cells before further processing. | Create a summary of populated cells to drive conditional formatting or reporting. | Skip empty rows when exporting worksheet data to CSV, JSON, or other formats. | Calculate occupancy of a dynamic range for data‑quality checks.
// AI Prompts: Write a reusable C# method that takes a Worksheet object and a range address string and returns the count of non‑empty cells using Aspose.Cells. | Show how to modify the loop so cells containing only whitespace are treated as empty while counting. | Explain how to use Cells.MaxDataRow and Cells.MaxDataColumn to determine the used range automatically and then count non‑empty cells. | Provide an example that logs the count and saves the workbook with a custom filename.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A concise C# example that creates a workbook, fills selected cells, defines a range (e.g., "A1:C5"), iterates the Aspose.Cells.Range collection, checks each Cell.Value for null or empty string, increments a counter for populated cells, writes the total to the console, and optionally saves the file.
    class CountNonEmptyCellsInRange
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate some sample data (mix of empty and non‑empty cells)
                cells["A1"].PutValue("Hello");
                cells["B2"].PutValue(123);
                cells["C3"].PutValue("World");
                // A2, B1, B3, etc. remain empty

                // Define the range to examine
                string rangeAddress = "A1:C5";

                // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
                Aspose.Cells.Range range = cells.CreateRange(rangeAddress);

                // Count non‑empty cells in the range
                int nonEmptyCount = 0;
                foreach (Cell cell in range)
                {
                    // Cell.Value is null for empty cells; also check for empty string after conversion
                    if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        nonEmptyCount++;
                    }
                }

                // Output the result
                Console.WriteLine($"Non‑empty cells in range {rangeAddress}: {nonEmptyCount}");

                // Save the workbook (optional, just to demonstrate lifecycle compliance)
                workbook.Save("CountNonEmptyCellsInRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
