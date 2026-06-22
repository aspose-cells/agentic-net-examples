using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Load the workbook from disk
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (you can change the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells C6:E7.
        // C = column index 2, 6 = row index 5 (zero‑based).
        // Total rows = 2 (rows 6 and 7), total columns = 3 (C, D, E).
        worksheet.Cells.Merge(5, 2, 2, 3);

        // Apply bold font to the merged cell (reference the upper‑left cell C6).
        Style style = worksheet.Cells[5, 2].GetStyle();
        style.Font.IsBold = true;
        worksheet.Cells[5, 2].SetStyle(style);

        // Save the modified workbook as XLSX.
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}