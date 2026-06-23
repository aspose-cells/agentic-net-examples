using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells P2:P5 (zero‑based indices: row 1, column 15, 4 rows, 1 column)
            worksheet.Cells.Merge(1, 15, 4, 1);

            // Set wrap text on the merged cell (top‑left cell of the merged range)
            Style mergedStyle = worksheet.Cells[1, 15].GetStyle();
            mergedStyle.IsTextWrapped = true;
            worksheet.Cells[1, 15].SetStyle(mergedStyle);

            // Apply an auto‑filter to the merged range (optional header row can be added if needed)
            worksheet.AutoFilter.Range = "P2:P5";
            worksheet.AutoFilter.Refresh();

            // Save the workbook as XLSX
            workbook.Save("MergedWrapAutoFilter.xlsx", SaveFormat.Xlsx);
        }
    }
}