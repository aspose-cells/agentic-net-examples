// Title: Aspose.Cells for .NET – Build an Excel Table with a Merged Title Row Across Multiple Columns
// Description: Demonstrates creating a Workbook, merging cells A1:D1 into a centered bold title, adding a styled header row, filling sample data from a 2‑D array, auto‑fitting columns, and saving as MergedHeaderTable.xlsx using C#.
// Keywords: Aspose.Cells | .NET | C# | merge cells | merged header | Excel table | auto fit columns | style title | sample data array | Workbook | Worksheet
// Common Searches: Aspose.Cells merge cells C# | how to create merged header row in Excel with Aspose.Cells | style merged title Aspose.Cells .NET | auto fit columns after data Aspose.Cells | add header row after merged title Aspose.Cells | populate worksheet from 2D array Aspose.Cells
// Developer Intent: Create an Excel worksheet with a merged, formatted title row, a bold header row, and populated data using Aspose.Cells for .NET.
// Use Cases: Generate a sales report with a centered title spanning columns A‑D. | Design an invoice sheet where the heading covers the full table width. | Build a dashboard worksheet that requires a merged title above metric tables. | Prepare a project status sheet with a bold merged header for easy printing.
// AI Prompts: Show C# code to merge cells A1:D1 and apply center‑aligned bold formatting with Aspose.Cells. | Provide an example that adds a styled header row after a merged title and auto‑fits columns in an Aspose.Cells workbook. | Explain how to fill a worksheet from a two‑dimensional array and set a background color for the header row using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsTableWithMergedHeader
{
    // Demonstrates creating a Workbook, merging cells A1:D1 into a centered bold title, adding a styled header row, filling sample data from a 2‑D array, auto‑fitting columns, and saving as MergedHeaderTable.xlsx using C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Create a merged title that spans columns A to D
            // -------------------------------------------------
            // Merge cells A1:D1 (zero‑based indices: row 0, column 0, 1 row, 4 columns)
            worksheet.Cells.Merge(0, 0, 1, 4);
            // Set the title text in the merged cell (reference the upper‑left cell)
            worksheet.Cells[0, 0].PutValue("Sales Report 2024");

            // Optional: style the title (center alignment, bold, larger font)
            Style titleStyle = worksheet.Cells[0, 0].GetStyle();
            titleStyle.HorizontalAlignment = TextAlignmentType.Center;
            titleStyle.VerticalAlignment = TextAlignmentType.Center;
            titleStyle.Font.IsBold = true;
            titleStyle.Font.Size = 14;
            worksheet.Cells[0, 0].SetStyle(titleStyle);

            // -------------------------------------------------
            // Add header row (first row after the title)
            // -------------------------------------------------
            // Header values in cells A2, B2, C2, D2 (row index 1)
            worksheet.Cells[1, 0].PutValue("Product");
            worksheet.Cells[1, 1].PutValue("Region");
            worksheet.Cells[1, 2].PutValue("Units Sold");
            worksheet.Cells[1, 3].PutValue("Revenue");

            // Optional: style the header row
            Style headerStyle = worksheet.Cells[1, 0].GetStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            for (int col = 0; col < 4; col++)
            {
                worksheet.Cells[1, col].SetStyle(headerStyle);
            }

            // -------------------------------------------------
            // Add some sample data rows
            // -------------------------------------------------
            string[,] data = new string[,]
            {
                { "Laptop", "North", "120", "150000" },
                { "Smartphone", "South", "300", "210000" },
                { "Tablet", "East", "80", "56000" }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                int row = 2 + i; // Data starts at row index 2 (Excel row 3)
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    worksheet.Cells[row, col].PutValue(data[i, col]);
                }
            }

            // -------------------------------------------------
            // Auto‑fit columns for better appearance
            // -------------------------------------------------
            worksheet.AutoFitColumns();

            // Save the workbook to an XLSX file
            workbook.Save("MergedHeaderTable.xlsx");
        }
    }
}
