// Title: Iterate Populated Cells in Aspose.Cells (C#) with MaxDataRow & MaxDataColumn
// Description: This C# example creates a workbook, inserts sample data with an intentional blank row, retrieves the worksheet's MaxDataRow and MaxDataColumn a single time, and loops through the exact data region, outputting only cells that hold a value before saving the file. The technique reduces API calls and skips empty rows and columns for optimal performance.
// Keywords: Aspose.Cells | C# | MaxDataRow | MaxDataColumn | iterate populated cells | skip empty rows | Excel performance tip | efficient cell iteration | Aspose.Cells API usage | Excel data extraction .NET
// Common Searches: How to loop through only non‑empty cells in Aspose.Cells C# | Using MaxDataRow and MaxDataColumn to avoid blank rows in Excel | Best practice for iterating a data range with Aspose.Cells | Performance tip: cache MaxDataRow in Aspose.Cells | Aspose.Cells skip empty rows while reading worksheet
// Developer Intent: I want to traverse only the cells that contain data in an Excel worksheet, ignoring blank rows and columns, by leveraging MaxDataRow and MaxDataColumn in Aspose.Cells for .NET.
// Use Cases: Display values from a sheet while omitting empty rows for clean reporting. | Copy only rows with data to another worksheet or external database to reduce processing time. | Apply formatting, formulas, or calculations exclusively to populated cells in large workbooks.
// AI Prompts: Write C# code that uses Aspose.Cells to iterate over the populated area of a worksheet with MaxDataRow and MaxDataColumn, skipping empty cells. | Show an example of copying only non‑blank rows from one worksheet to another using Aspose.Cells. | Explain why caching MaxDataRow and MaxDataColumn improves iteration speed in large Excel files with Aspose.Cells.

using System;
using Aspose.Cells;

// This C# example creates a workbook, inserts sample data with an intentional blank row, retrieves the worksheet's MaxDataRow and MaxDataColumn a single time, and loops through the exact data region, outputting only cells that hold a value before saving the file. The technique reduces API calls and skips empty rows and columns for optimal performance.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some sample data (including a deliberately empty row)
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Score");
        cells["A2"].PutValue("Alice");
        cells["B2"].PutValue(85);
        cells["A3"].PutValue("Bob");
        cells["B3"].PutValue(92);
        // Row 4 (index 3) left empty
        cells["A5"].PutValue("Charlie");
        cells["B5"].PutValue(78);

        // Retrieve the limits only once – calling MaxDataRow/MaxDataColumn repeatedly is expensive
        int maxRow = cells.MaxDataRow;       // zero‑based index of the last row that contains data
        int maxCol = cells.MaxDataColumn;    // zero‑based index of the last column that contains data

        // Iterate through the populated area only
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                // Skip null cells or cells with no value
                if (cell != null && cell.Value != null && !string.IsNullOrEmpty(cell.StringValue))
                {
                    Console.WriteLine($"Cell {cell.Name} = {cell.Value}");
                }
            }
        }

        // Save the workbook (uses the provided save rule)
        workbook.Save("IteratePopulatedCells.xlsx");
    }
}
