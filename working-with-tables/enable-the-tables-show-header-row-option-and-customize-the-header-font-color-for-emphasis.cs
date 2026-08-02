// Title: Show Table Header Row and Apply Red Bold Font with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a ListObject (Excel table) over a data range, enables the header row with ShowHeaderRow, defines a red bold style, applies it to each header cell, and saves the file as TableHeaderDemo.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | Excel table header | ShowHeaderRow | ListObject | header font color | red bold style | custom table header | Excel export | table formatting
// Common Searches: Aspose.Cells show header row | How to style Excel table header with Aspose.Cells | Set ListObject ShowHeaderRow property C# | Change font color of table header Aspose.Cells | Apply bold style to Excel table header using .NET
// Developer Intent: Enable the table’s header row and format its font color and weight.
// Use Cases: Generate a sales report where column titles are highlighted in red bold for quick scanning. | Create a financial worksheet that requires visible, styled headers to meet corporate branding. | Export data to Excel from an application and ensure the table headers stand out for end‑users.
// AI Prompts: Write C# code with Aspose.Cells that adds a ListObject, sets ShowHeaderRow to true, and formats the header font to blue and italic. | Provide an example that loops through a table’s header cells and applies a style with background shading and borders using Aspose.Cells. | Explain how to toggle ShowHeaderRow at runtime and refresh the header style dynamically in a .NET application.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableHeaderDemo
{
    // This example creates a workbook, adds a ListObject (Excel table) over a data range, enables the header row with ShowHeaderRow, defines a red bold style, applies it to each header cell, and saves the file as TableHeaderDemo.xlsx using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (including header row)
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1.20);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(0.80);
            worksheet.Cells["A4"].PutValue("Cherry");
            worksheet.Cells["B4"].PutValue(2.00);

            // Add a ListObject (table) covering the data range
            // Parameters: first row, first column, last row, last column, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Ensure the header row is visible
            table.ShowHeaderRow = true;

            // Create a style for the header cells (e.g., red font, bold)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Color = Color.Red;
            headerStyle.Font.IsBold = true;

            // Apply the style to each header cell in the table
            int startRow = table.StartRow;          // Header row index
            int startCol = table.StartColumn;       // First column index
            int endCol = table.EndColumn;           // Last column index

            for (int col = startCol; col <= endCol; col++)
            {
                Cell headerCell = worksheet.Cells[startRow, col];
                headerCell.SetStyle(headerStyle);
            }

            // Save the workbook
            workbook.Save("TableHeaderDemo.xlsx");
        }
    }
}
