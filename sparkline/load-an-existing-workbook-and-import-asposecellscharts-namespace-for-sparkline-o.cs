using System;
using Aspose.Cells;
using Aspose.Cells.Charts; // Namespace for sparkline operations

namespace SparklineDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing Excel file to be loaded
            string inputPath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (if not already present)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1); // Column A
                sheet.Cells[i, 1].PutValue((i + 1) * 2); // Column B
                sheet.Cells[i, 2].PutValue((i + 1) * 3); // Column C
                sheet.Cells[i, 3].PutValue((i + 1) * 4); // Column D
            }

            // Define the location range for the sparkline group (E1:E5)
            CellArea location = CellArea.CreateCellArea("E1", "E5");

            // Add a sparkline group of type Line, using data from A1:D5, placed at E1:E5
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D5", false, location);
            SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

            // Customize the sparkline group
            sparklineGroup.ShowHighPoint = true;
            sparklineGroup.ShowLowPoint = true;

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook loaded from '{inputPath}', sparkline added, and saved to '{outputPath}'.");
        }
    }
}