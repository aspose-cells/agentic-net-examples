using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class DynamicSparklineDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate initial data in column A (A1:A5)
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1); // values 1,2,3,4,5
        }

        // Add a sparkline group of type Line
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Use a formula‑based data range that expands with the number of non‑empty cells in column A
        // The formula "A1:INDEX(A:A, COUNTA(A:A))" creates a range from A1 to the last filled cell in column A
        string dynamicRange = "A1:INDEX(A:A, COUNTA(A:A))";

        // Add a sparkline that uses the dynamic range.
        // Place the sparkline in cell B1 (row 0, column 1)
        group.Sparklines.Add(dynamicRange, 0, 1);

        // Add more data to column A to demonstrate automatic update
        sheet.Cells[5, 0].PutValue(6); // A6
        sheet.Cells[6, 0].PutValue(7); // A7

        // Recalculate formulas so that the dynamic range reflects the new data
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("DynamicSparklineDemo.xlsx");
    }
}