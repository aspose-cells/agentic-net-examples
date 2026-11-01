// Title: Create a merged title row spanning columns A‑D with bold centered text in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge cells A1 to D1 for a table title | set bold centered title across multiple columns in Excel with Aspose.Cells .NET | create header row that spans columns in an Aspose.Cells workbook using C# | apply style to merged cells in Aspose.Cells C# example | auto‑fit columns after merging cells with Aspose.Cells for .NET
// Tags: merge cells Aspose.Cells .NET | styled merged title Excel Aspose | header row spanning columns Aspose.Cells | auto fit columns Aspose.Cells workbook | populate sample data Aspose.Cells worksheet

using System;
using Aspose.Cells;

namespace AsposeCellsHeaderTableExample
{
    // The example creates a new workbook, merges cells A1:D1 to form a centered bold title, adds a formatted header row, fills sample data rows, auto‑fits the columns, and saves the file as HeaderTable.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Create a merged title cell that spans columns A to D.
            // -------------------------------------------------
            // Merge cells A1:D1 (zero‑based indexes: row 0, columns 0‑3).
            sheet.Cells.Merge(0, 0, 1, 4);
            // Set the title text.
            sheet.Cells["A1"].PutValue("Sales Report 2024");
            // Apply simple formatting to the title.
            Style titleStyle = workbook.CreateStyle();
            titleStyle.Font.IsBold = true;
            titleStyle.Font.Size = 14;
            titleStyle.HorizontalAlignment = TextAlignmentType.Center;
            sheet.Cells["A1"].SetStyle(titleStyle);

            // -------------------------------------------------
            // 2. Add header row (row 2) with column names.
            // -------------------------------------------------
            int headerRow = 1; // zero‑based index for the second row.
            sheet.Cells[headerRow, 0].PutValue("Product");
            sheet.Cells[headerRow, 1].PutValue("Region");
            sheet.Cells[headerRow, 2].PutValue("Quantity");
            sheet.Cells[headerRow, 3].PutValue("Revenue");

            // Apply header style.
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            for (int col = 0; col < 4; col++)
            {
                sheet.Cells[headerRow, col].SetStyle(headerStyle);
            }

            // -------------------------------------------------
            // 3. Populate some sample data rows.
            // -------------------------------------------------
            object[,] data = new object[,]
            {
                { "Widget A", "North", 120, 2400 },
                { "Widget B", "South", 85, 1700 },
                { "Widget C", "East",  60, 1800 },
                { "Widget D", "West",  95, 2850 }
            };

            int startDataRow = headerRow + 1;
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    sheet.Cells[startDataRow + i, j].PutValue(data[i, j]);
                }
            }

            // -------------------------------------------------
            // 4. Auto‑fit columns for better appearance.
            // -------------------------------------------------
            sheet.AutoFitColumns();

            // Save the workbook to a file.
            workbook.Save("HeaderTable.xlsx");
        }
    }
}
