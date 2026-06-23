using System;
using Aspose.Cells;

class MergeAndExportCsv
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells F4:G5 (zero‑based indices: row 3, column 5, 2 rows, 2 columns)
        cells.Merge(3, 5, 2, 2);

        // Put a sample value into the merged cell (optional, helps see the formula result)
        cells["F4"].PutValue(10);

        // Insert a formula in another cell that references the merged cell (e.g., H4 = F4 * 2)
        cells["H4"].Formula = "=F4*2";

        // Calculate formulas so the result is stored before exporting
        workbook.CalculateFormula();

        // Export the worksheet to CSV
        string outputPath = "output.csv";
        workbook.Save(outputPath, SaveFormat.Csv);
    }
}