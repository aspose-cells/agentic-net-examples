using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsMergeStyledExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided Workbook constructor)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells L5:M7
            // L = column index 11, row 5 = row index 4, total rows = 3, total columns = 2
            worksheet.Cells.Merge(4, 11, 3, 2);

            // Create a custom style (uses Workbook.CreateStyle rule)
            Style customStyle = workbook.CreateStyle();

            // Set italic font
            customStyle.Font.IsItalic = true;

            // Set gray fill background
            customStyle.ForegroundColor = Color.Gray;
            customStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the merged cell (upper‑left cell of the merged range)
            Cell mergedCell = worksheet.Cells[4, 11];
            mergedCell.SetStyle(customStyle);

            // Save the workbook (uses the standard Save method)
            workbook.Save("MergedStyled.xlsx", SaveFormat.Xlsx);
        }
    }
}