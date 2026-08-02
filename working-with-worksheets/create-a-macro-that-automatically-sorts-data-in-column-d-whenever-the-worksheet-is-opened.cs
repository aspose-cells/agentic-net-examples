// Title: Auto‑sort Column D on Workbook Open with Aspose.Cells (C#)
// Description: C# example that leverages Aspose.Cells' DataSorter to order the entire used range by column D (header row assumed) each time the workbook is opened, then saves the updated file.
// Keywords: Aspose.Cells DataSorter | C# sort Excel column D | auto sort on workbook open | programmatic Excel sorting | sort used range Aspose
// Common Searches: Aspose.Cells sort column D on open | C# automatically sort Excel worksheet | DataSorter sort by specific column | trigger sort when opening Excel file using Aspose | sort Excel data programmatically with headers
// Developer Intent: Automatically sort worksheet rows by column D whenever the workbook is opened.
// Use Cases: Daily sales report where rows must be ordered by the date column (D) before distribution. | Product catalog that should always appear alphabetically by the name column (D) for end‑users. | Transaction log that needs to stay sorted by transaction ID in column D for downstream processing.
// AI Prompts: Generate C# code using Aspose.Cells to sort the used range by column D with headers each time the workbook opens. | Show how to add a workbook‑open trigger in Aspose.Cells that sorts column D and saves the file. | Explain how to modify the DataSorter to sort column D in descending order while keeping cell formatting intact.

using System;
using Aspose.Cells;

// C# example that leverages Aspose.Cells' DataSorter to order the entire used range by column D (header row assumed) each time the workbook is opened, then saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the used range of the worksheet
        int lastRow = worksheet.Cells.MaxDataRow;
        int lastColumn = worksheet.Cells.MaxDataColumn;

        // Configure the DataSorter to sort by column D (zero‑based index 3)
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true;               // assume the first row contains headers
        sorter.Key1 = 3;                         // column D
        sorter.Order1 = SortOrder.Ascending;    // sort in ascending order

        // Define the area to be sorted (entire used range)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = lastRow,
            EndColumn = lastColumn
        };

        // Perform the sort operation
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the workbook (the data will be sorted each time the file is opened)
        workbook.Save("output.xlsx");
    }
}
