using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class AutoFitRowExample
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Modify a cell value (e.g., B2)
            worksheet.Cells["B2"].PutValue("This is a new, longer piece of text that should cause the row height to increase.");

            // Auto‑fit the row that contains the modified cell (row index is zero‑based, so B2 is row 1)
            worksheet.AutoFitRow(1);

            // Save the updated workbook (replace with your desired output path)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}