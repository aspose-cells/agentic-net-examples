// Title: Sum Sales Column with Cells.Subtotal Using Header‑Based Index in Aspose.Cells for .NET
// Description: Creates a workbook, adds a header row containing "Sales", scans the header to obtain the column index, defines a CellArea for the data range, and applies Worksheet.Cells.Subtotal to group by Region and sum the Sales values before saving the file.
// Keywords: Aspose.Cells | Cells.Subtotal | C# | dynamic column index | header lookup | sum sales column | group by region | ConsolidationFunction.Sum | Excel subtotal automation
// Common Searches: Aspose.Cells subtotal by header name | C# find column index from header and sum | Worksheet.Cells.Subtotal dynamic column | group rows by region and subtotal sales in Aspose.Cells | how to use Cells.Subtotal with runtime column index
// Developer Intent: Create a regional sales report that automatically locates the "Sales" column by its header and inserts subtotal rows that sum those values.
// Use Cases: Generating sales summaries where the column order may change. | Automating financial spreadsheets that require subtotal rows for a metric identified by its header. | Building Excel exports that group data by a key field and provide dynamic subtotals for numeric columns.
// AI Prompts: Write C# code using Aspose.Cells that locates a column named "Sales" from the header row and applies Worksheet.Cells.Subtotal to sum it grouped by the first column. | Show an example of Cells.Subtotal with ConsolidationFunction.Sum where the target column index is determined at runtime from a header string. | Explain step‑by‑step how to find a column by its header and add subtotal rows in Aspose.Cells for .NET.

using Aspose.Cells;
using System;

// Creates a workbook, adds a header row containing "Sales", scans the header to obtain the column index, defines a CellArea for the data range, and applies Worksheet.Cells.Subtotal to group by Region and sum the Sales values before saving the file.
class SubtotalBySalesColumn
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data with a header that includes "Sales"
        cells["A1"].PutValue("Region");
        cells["B1"].PutValue("Product");
        cells["C1"].PutValue("Sales");

        object[,] data = new object[,]
        {
            { "North", "Widget", 5000 },
            { "North", "Gadget", 3000 },
            { "South", "Widget", 6000 },
            { "South", "Gadget", 4000 },
            { "West",  "Widget", 4500 }
        };

        for (int row = 0; row < data.GetLength(0); row++)
        {
            for (int col = 0; col < data.GetLength(1); col++)
            {
                cells[row + 1, col].PutValue(data[row, col]);
            }
        }

        // Determine the zero‑based column index of the header named "Sales"
        int salesColumnIndex = -1;
        int headerRow = 0;
        int lastColumn = cells.MaxDataColumn;

        for (int col = 0; col <= lastColumn; col++)
        {
            if (cells[headerRow, col].StringValue == "Sales")
            {
                salesColumnIndex = col;
                break;
            }
        }

        if (salesColumnIndex == -1)
        {
            Console.WriteLine("Column \"Sales\" not found.");
            return;
        }

        // Define the cell area that contains the data (including the header row)
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = cells.MaxDataRow,
            EndColumn = cells.MaxDataColumn
        };

        // Apply subtotal:
        // - Group by the first column (Region, index 0)
        // - Use SUM function
        // - Add subtotal to the Sales column identified above
        worksheet.Cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { salesColumnIndex });

        // Save the workbook
        workbook.Save("SubtotalBySales.xlsx");
    }
}
