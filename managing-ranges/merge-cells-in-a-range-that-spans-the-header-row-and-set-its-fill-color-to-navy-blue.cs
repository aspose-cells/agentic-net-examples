using System;
using System.Drawing;
using Aspose.Cells;

namespace MergeHeaderExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the header range (first row, columns A to E)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 1;         // only the header row
            int totalColumns = 5;      // columns A‑E

            // Merge the header range into a single cell
            cells.Merge(firstRow, firstColumn, totalRows, totalColumns);

            // Create a style with navy blue fill
            Style headerStyle = cells[firstRow, firstColumn].GetStyle();
            headerStyle.Pattern = BackgroundType.Solid;   // enable solid fill
            headerStyle.ForegroundColor = Color.Navy;     // set fill color to navy blue
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.VerticalAlignment = TextAlignmentType.Center;
            headerStyle.Font.IsBold = true;

            // Apply the style to the merged cell (upper‑left cell of the range)
            cells[firstRow, firstColumn].SetStyle(headerStyle);

            // Optionally set a header value
            cells[firstRow, firstColumn].PutValue("Report Header");

            // Save the workbook
            workbook.Save("MergedHeader.xlsx");
        }
    }
}