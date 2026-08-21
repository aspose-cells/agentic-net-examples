// Title: C# – Create an Excel table with a merged title row using Aspose.Cells for .NET
// Description: A complete C# example that creates a new workbook, merges cells A1:D1 for a centered title, adds a header row and sample data, defines a ListObject table, auto‑fits columns, and saves the file as TableWithMergedTitle.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# Excel merge cells | merged title row | ListObject table | auto fit columns | save workbook .xlsx | Aspose.Cells example | GitHub sample | Excel table creation | Aspose.Cells .NET
// Common Searches: Aspose.Cells merge cells for title row C# | Create ListObject table after merged header Aspose.Cells | Save Excel file with merged title using Aspose.Cells .NET | C# example for merged title and table in Excel | How to auto‑fit columns in Aspose.Cells C#
// Developer Intent: Generate an Excel worksheet that features a merged title spanning the table width, followed by a formatted ListObject table, using Aspose.Cells in a .NET C# project.
// Use Cases: Produce monthly sales reports with a full‑width title and a data table that can be filled programmatically. | Build reusable Excel templates that include a centered heading and a structured table for downstream analysis. | Export database query results to Excel with a merged header and a styled ListObject for easy filtering and sorting.
// AI Prompts: Write C# code with Aspose.Cells to create a worksheet, merge cells A1:D1 for a title, add a header row and sample rows, convert the range into a ListObject table, auto‑fit columns, and save as an .xlsx file. | Show a complete Aspose.Cells .NET example that demonstrates merging cells for a title row, creating a table with headers, setting the table’s display name, and adjusting column widths. | Explain how to apply alignment and styling to a merged title row after merging cells using Aspose.Cells in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// A complete C# example that creates a new workbook, merges cells A1:D1 for a centered title, adds a header row and sample data, defines a ListObject table, auto‑fits columns, and saves the file as TableWithMergedTitle.xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ----- Title row (merged across 4 columns) -----
        // Put the title text in the upper‑left cell of the range
        worksheet.Cells[0, 0].PutValue("Sales Report 2023");
        // Merge cells A1:D1 (row 0, column 0, 1 row, 4 columns)
        worksheet.Cells.Merge(0, 0, 1, 4);

        // ----- Table header row (row 2) -----
        worksheet.Cells[1, 0].PutValue("Product");
        worksheet.Cells[1, 1].PutValue("Region");
        worksheet.Cells[1, 2].PutValue("Units Sold");
        worksheet.Cells[1, 3].PutValue("Revenue");

        // ----- Sample data rows -----
        worksheet.Cells[2, 0].PutValue("Apple");
        worksheet.Cells[2, 1].PutValue("North");
        worksheet.Cells[2, 2].PutValue(120);
        worksheet.Cells[2, 3].PutValue(2400);

        worksheet.Cells[3, 0].PutValue("Banana");
        worksheet.Cells[3, 1].PutValue("South");
        worksheet.Cells[3, 2].PutValue(85);
        worksheet.Cells[3, 3].PutValue(1275);

        // ----- Create a ListObject (table) -----
        // Table starts at the header row (row index 1) and includes the data rows
        int firstRow = 1;          // zero‑based index for row 2
        int firstColumn = 0;       // column A
        int totalRows = 3;         // header + 2 data rows
        int totalColumns = 4;      // columns A‑D

        int tableIndex = worksheet.ListObjects.Add(
            firstRow,
            firstColumn,
            firstRow + totalRows - 1,
            firstColumn + totalColumns - 1,
            true); // true => has headers

        ListObject table = worksheet.ListObjects[tableIndex];
        table.DisplayName = "SalesData";

        // Adjust column widths for readability
        worksheet.AutoFitColumns();

        // Save the workbook
        workbook.Save("TableWithMergedTitle.xlsx");
    }
}
