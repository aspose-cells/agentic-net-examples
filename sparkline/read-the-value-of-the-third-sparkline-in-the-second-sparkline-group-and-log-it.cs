using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineReadDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two sparkline groups
            // Group 1 data (row 0)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(2);
            // Group 2 data (row 1)
            sheet.Cells["A2"].PutValue(7);
            sheet.Cells["B2"].PutValue(1);
            sheet.Cells["C2"].PutValue(4);
            sheet.Cells["D2"].PutValue(6);

            // Define location ranges for the sparkline groups (single cells)
            CellArea locationGroup1 = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };
            CellArea locationGroup2 = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 5,
                EndColumn = 5
            };

            // Add first sparkline group (adds first sparkline automatically)
            int groupIdx1 = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, locationGroup1);
            SparklineGroup group1 = sheet.SparklineGroups[groupIdx1];
            // Add additional sparklines to first group (rows 1 and 2)
            group1.Sparklines.Add(sheet.Name + "!A1:D1", 1, 4);
            group1.Sparklines.Add(sheet.Name + "!A1:D1", 2, 4);

            // Add second sparkline group
            int groupIdx2 = sheet.SparklineGroups.Add(SparklineType.Line, "A2:D2", false, locationGroup2);
            SparklineGroup group2 = sheet.SparklineGroups[groupIdx2];
            // Add additional sparklines to second group (rows 1 and 2)
            group2.Sparklines.Add(sheet.Name + "!A2:D2", 1, 5);
            group2.Sparklines.Add(sheet.Name + "!A2:D2", 2, 5);

            // Read the DataRange of the third sparkline in the second group
            SparklineGroup targetGroup = sheet.SparklineGroups[1];
            Sparkline targetSparkline = targetGroup.Sparklines[2];
            Console.WriteLine("DataRange of the third sparkline in the second group: " + targetSparkline.DataRange);

            // Save the workbook
            workbook.Save("SparklineReadDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}