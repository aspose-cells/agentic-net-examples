using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

class SparklineDynamicDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate initial data in column A (A1:A5)
        for (int i = 0; i < 5; i++)
        {
            worksheet.Cells[i, 0].PutValue(i + 1);
        }

        // Add a sparkline group of type Line
        int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

        // Define a formula‑based data range that expands automatically.
        // OFFSET starts at A1, height is the count of non‑empty cells in column A.
        string dynamicDataRange = "OFFSET(A1,0,0,COUNTA(A:A),1)";

        // Add a sparkline at cell B1 (row 0, column 1) using the dynamic range
        sparklineGroup.Sparklines.Add(dynamicDataRange, 0, 1);

        // Save the workbook with the initial sparkline
        workbook.Save("DynamicSparkline.xlsx");

        // ----- Demonstrate automatic update -----
        // Append more values to column A (A6 and A7)
        worksheet.Cells[5, 0].PutValue(6);
        worksheet.Cells[6, 0].PutValue(7);

        // Recalculate formulas to ensure any dependent calculations are refreshed
        workbook.CalculateFormula();

        // Save the workbook after data change; the sparkline reflects the new range automatically
        workbook.Save("DynamicSparkline_Updated.xlsx");
    }
}