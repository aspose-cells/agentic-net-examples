using System;
using Aspose.Cells;

namespace AsposeCellsMergeAndRotate
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells Z4:AA4 (zero‑based indices: row 3, column 25, 1 row, 2 columns)
            worksheet.Cells.Merge(3, 25, 1, 2);

            // Create a style with a 45‑degree text rotation
            Style rotationStyle = workbook.CreateStyle();
            rotationStyle.RotationAngle = 45;

            // Enable the rotation flag so the style is applied
            StyleFlag flag = new StyleFlag();
            flag.Rotation = true;

            // Apply the style to the merged cell (upper‑left cell of the range)
            worksheet.Cells["Z4"].SetStyle(rotationStyle, flag);

            // Save the modified workbook as XLSX
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}