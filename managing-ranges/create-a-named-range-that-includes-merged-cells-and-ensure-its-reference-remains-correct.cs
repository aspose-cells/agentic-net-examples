// Title: Aspose.Cells .NET: Create a Named Range that Includes Merged Cells
// Description: Demonstrates how to merge cells A1:B2, assign a value, add a named range called "MergedArea" with an absolute A1 RefersTo formula, verify the range address and value, and save the workbook as NamedRangeWithMergedCells.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells named range merged cells | C# Aspose.Cells create named range | merged cells reference Aspose | absolute A1 notation Aspose.Cells | verify named range address .NET
// Common Searches: Aspose.Cells add named range for merged cells | C# create named range covering A1:B2 merged area | how to keep named range reference after merging cells Aspose | retrieve values from a named range that includes merged cells
// Developer Intent: Add a named range that spans merged cells and ensure its RefersTo formula remains accurate.
// Use Cases: Define a reusable range for a merged header that can be used in formulas or charts. | Programmatically read or modify data inside a merged block via a named range. | Export a workbook with a correctly referenced named range for downstream reporting.
// AI Prompts: Write C# code with Aspose.Cells to merge A1:B2, create a named range for the merged area, and print its address. | Explain how to set the RefersTo property using absolute A1 notation for a merged range in Aspose.Cells. | Show how to retrieve a named range that contains merged cells and access its first cell value.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeWithMergedCells
{
    // Demonstrates how to merge cells A1:B2, assign a value, add a named range called "MergedArea" with an absolute A1 RefersTo formula, verify the range address and value, and save the workbook as NamedRangeWithMergedCells.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (default name is "Sheet1")
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Merge cells A1:B2 (zero‑based indices: row 0, column 0, 2 rows, 2 columns)
                cells.Merge(0, 0, 2, 2);

                // Put a value into the merged cell (upper‑left cell of the range)
                cells[0, 0].PutValue("Merged Area");

                // Add a named range that covers the merged cells
                int nameIndex = workbook.Worksheets.Names.Add("MergedArea");
                Name mergedName = workbook.Worksheets.Names[nameIndex];
                // RefersTo must start with '=' and use absolute A1 notation
                mergedName.RefersTo = $"={sheet.Name}!$A$1:$B$2";

                // Retrieve the range via the name to verify the reference is correct
                Aspose.Cells.Range namedRange = mergedName.GetRange();
                Console.WriteLine($"Named range address: {namedRange.Address}");
                Console.WriteLine($"First cell of the range value: {namedRange[0, 0].StringValue}");

                // Save the workbook
                string outputPath = "NamedRangeWithMergedCells.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
