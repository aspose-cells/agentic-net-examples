using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class RemoveThirdChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the charts
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value1");
        sheet.Cells["C1"].PutValue("Value2");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["C4"].PutValue(35);

        // Add three charts to the worksheet
        int chartIndex1 = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
        Chart chart1 = sheet.Charts[chartIndex1];
        chart1.NSeries.Add("B2:B4", true);
        chart1.NSeries.CategoryData = "A2:A4";

        int chartIndex2 = sheet.Charts.Add(ChartType.Line, 16, 0, 26, 8);
        Chart chart2 = sheet.Charts[chartIndex2];
        chart2.NSeries.Add("C2:C4", true);
        chart2.NSeries.CategoryData = "A2:A4";

        int chartIndex3 = sheet.Charts.Add(ChartType.Pie, 27, 0, 37, 8);
        Chart chart3 = sheet.Charts[chartIndex3];
        chart3.NSeries.Add("B2:B4", true);
        chart3.NSeries.CategoryData = "A2:A4";

        // Display chart count before removal
        Console.WriteLine("Chart count before removal: " + sheet.Charts.Count);

        // Remove the third chart (zero‑based index 2)
        sheet.Charts.RemoveAt(2);

        // Display chart count after removal
        Console.WriteLine("Chart count after removal: " + sheet.Charts.Count);

        // Save the workbook
        workbook.Save("RemovedThirdChart.xlsx");
    }
}