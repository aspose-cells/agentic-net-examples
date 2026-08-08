// Title: Aspose.Cells .NET: Iterate Only Populated Cells with MaxDataRow & MaxDataColumn
// Description: Creates a workbook, inserts data with gaps, obtains the zero‑based last‑filled row and column via MaxDataRow/MaxDataColumn, and loops through that range. The example processes only cells that contain a value, converts string entries to uppercase, and saves the modified file.
// Keywords: Aspose.Cells | MaxDataRow | MaxDataColumn | C# worksheet iteration | skip empty rows | skip empty columns | populate range loop | convert string to uppercase | save workbook | GitHub example
// Common Searches: Aspose.Cells iterate only filled cells | MaxDataRow vs MaxRow Aspose.Cells | skip empty rows and columns C# Aspose.Cells | process non‑empty cells Aspose.Cells .NET | example using MaxDataColumn in C#
// Developer Intent: Loop through the actual data region of a worksheet while ignoring blank rows and columns.
// Use Cases: Transform all textual data in the used range to uppercase. | Calculate aggregates (sum, average) on numeric cells without scanning empty cells. | Apply in‑place formatting or validation only to cells that contain values. | Export a cleaned data set after removing gaps.
// AI Prompts: Provide C# code that uses MaxDataRow and MaxDataColumn to iterate only non‑empty cells and convert strings to uppercase with Aspose.Cells. | Generate a .NET example that reads a worksheet, skips blank rows/columns, processes each value, and saves the workbook. | Explain the difference between MaxDataRow/MaxDataColumn and MaxRow/MaxColumn in Aspose.Cells and when each should be used.

using System;
using Aspose.Cells;

namespace AsposeCellsMaxDataIteration
{
    // Creates a workbook, inserts data with gaps, obtains the zero‑based last‑filled row and column via MaxDataRow/MaxDataColumn, and loops through that range. The example processes only cells that contain a value, converts string entries to uppercase, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data with gaps (empty rows/columns)
            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["A2"].PutValue("Row1Col1");
            cells["B2"].PutValue(100);
            // Row 3 is intentionally left empty
            cells["A4"].PutValue("Row4Col1");
            cells["B4"].PutValue(400);
            // Add data in a later column to test MaxDataColumn
            cells["D2"].PutValue("ExtraColumn");

            // Retrieve the maximum populated row and column indices
            int maxRow = cells.MaxDataRow;       // zero‑based index of last row containing data
            int maxCol = cells.MaxDataColumn;    // zero‑based index of last column containing data

            Console.WriteLine($"Iterating rows 0..{maxRow}, columns 0..{maxCol}");

            // Loop through only the populated area
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    // Process only cells that actually contain a value
                    if (cell.Value != null)
                    {
                        Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
                        // Example processing: convert string values to upper case
                        if (cell.Type == CellValueType.IsString)
                        {
                            cell.PutValue(cell.StringValue.ToUpper());
                        }
                    }
                }
            }

            // Save the workbook to demonstrate that changes were applied
            workbook.Save("ProcessedData.xlsx");
        }
    }
}
