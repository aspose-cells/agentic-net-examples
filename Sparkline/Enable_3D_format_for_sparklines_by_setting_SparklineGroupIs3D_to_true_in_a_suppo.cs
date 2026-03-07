using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklines3D
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(2);
            sheet.Cells["E1"].PutValue(7);

            // Define the location where the sparkline will be placed (cell F1)
            CellArea location = CellArea.CreateCellArea("F1", "F1");

            // Add a sparkline group of type Line, using the data range A1:E1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, location);
            SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

            // Optionally, add additional sparklines if needed
            sparklineGroup.Sparklines.Add(sheet.Name + "!A1:E1", 0, location.StartColumn);

            // Save the workbook in XLSX format
            workbook.Save("Sparklines3D.xlsx", SaveFormat.Xlsx);
        }
    }
}