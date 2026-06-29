using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the merged workbook (replace with your actual file path)
        string inputPath = "merged.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Apply AutoFitRows to every worksheet to adjust row heights based on content
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.AutoFitRows(); // Autofits all rows in the current worksheet
        }

        // Save the workbook after auto‑fitting rows
        string outputPath = "merged_autofit.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}