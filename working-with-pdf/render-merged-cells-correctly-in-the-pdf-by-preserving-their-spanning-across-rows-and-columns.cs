using System;
using Aspose.Cells;

namespace MergedCellsPdfDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Put some sample data
            cells["A1"].PutValue("Merged Header");
            cells["A2"].PutValue("Row 1");
            cells["A3"].PutValue("Row 2");
            cells["B2"].PutValue("Data 1");
            cells["B3"].PutValue("Data 2");

            // Merge cells A1:B1 (first row, spanning two columns)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            cells.Merge(0, 0, 1, 2);

            // Merge cells A2:A3 (first column, spanning two rows)
            cells.Merge(1, 0, 2, 1);

            // Apply simple styling to merged cells for better visibility
            Style style = cells["A1"].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            style.Font.IsBold = true;
            cells["A1"].SetStyle(style);

            // Auto-fit rows and columns to ensure merged cells are displayed correctly
            worksheet.AutoFitRows();
            worksheet.AutoFitColumns();

            // Save the workbook as PDF, preserving merged cell spans
            workbook.Save("MergedCellsOutput.pdf", SaveFormat.Pdf);
        }
    }
}