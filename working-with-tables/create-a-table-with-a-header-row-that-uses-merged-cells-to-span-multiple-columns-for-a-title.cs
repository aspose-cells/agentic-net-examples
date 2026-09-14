// Title: How to create an Excel worksheet with a merged, centered title row and styled table header using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that merges cells A1:D1 into a single title cell, centers and bolds the text, adds a bold centered header row, imports a two‑dimensional array of data, auto‑fits the columns, and saves the workbook as an .xlsx file. | Show how to define and apply a custom Style for a merged title cell and for header cells in an Aspose.Cells workbook, including horizontal/vertical alignment, font size, and bold settings.
// Common Searches: Aspose.Cells C# merge cells for a report title and style it | C# create Excel table with merged header row using Aspose.Cells | How to import a two‑dimensional array into a worksheet with Aspose.Cells .NET | AutoFitColumns after adding data with Aspose.Cells C# example | Set bold and centered style for header row in Aspose.Cells workbook
// Tags: merge cells A1:D1 Aspose.Cells | apply bold centered style Aspose.Cells | import two-dimensional array Aspose.Cells | auto-fit columns Aspose.Cells | create Excel table with header Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// // Generates an Excel .xlsx file using Aspose.Cells for .NET: merges cells A1:D1 into a centered bold title, adds a bold centered header row, imports sample data via a two‑dimensional array, auto‑fits columns, and saves as TableWithMergedHeader.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells for the title (A1:D1)
        sheet.Cells.Merge(0, 0, 1, 4); // row, column, rows to merge, columns to merge
        Cell titleCell = sheet.Cells[0, 0];
        titleCell.PutValue("Sales Report Q1 2024");

        // Style the title
        Style titleStyle = workbook.CreateStyle();
        titleStyle.HorizontalAlignment = TextAlignmentType.Center;
        titleStyle.VerticalAlignment = TextAlignmentType.Center;
        titleStyle.Font.IsBold = true;
        titleStyle.Font.Size = 14;
        titleCell.SetStyle(titleStyle);

        // Header row (row 2) for the table
        string[] headers = { "Product", "Region", "Units Sold", "Revenue" };
        for (int i = 0; i < headers.Length; i++)
        {
            Cell headerCell = sheet.Cells[1, i];
            headerCell.PutValue(headers[i]);

            // Header style
            Style headerStyle = workbook.CreateStyle();
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.Font.IsBold = true;
            headerCell.SetStyle(headerStyle);
        }

        // Sample data rows starting from row 3
        object[,] data = {
            { "Widget A", "North", 120, 3600 },
            { "Widget B", "South", 85, 2550 },
            { "Widget C", "East", 150, 4500 },
            { "Widget D", "West", 95, 2850 }
        };
        sheet.Cells.ImportTwoDimensionArray(data, 2, 0);

        // Auto-fit columns for better appearance
        sheet.AutoFitColumns();

        // Save the workbook
        workbook.Save("TableWithMergedHeader.xlsx");
    }
}
