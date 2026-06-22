using System;
using Aspose.Cells;

namespace AsposeCellsMergePhoneNumberDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells AB6:AC8
            // AB -> column index 27, AC -> column index 28
            // Row 6 -> row index 5, Row 8 -> row index 7
            // Total rows = 3 (rows 6,7,8), total columns = 2 (AB,AC)
            cells.Merge(firstRow: 5, firstColumn: 27, totalRows: 3, totalColumns: 2);

            // Apply phone number format "(###) ###‑####" to the merged cell
            // The merged cell is referenced by its upper‑left cell (AB6)
            Style style = cells[5, 27].GetStyle();
            style.Custom = "(###) ###-####";
            cells[5, 27].SetStyle(style);

            // Save the modified workbook
            string outputPath = "MergedPhoneNumber.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}