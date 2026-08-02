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
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add three charts to the worksheet
        int chartIndex1 = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
        int chartIndex2 = sheet.Charts.Add(ChartType.Line, 16, 0, 26, 8);
        int chartIndex3 = sheet.Charts.Add(ChartType.Pie, 27, 0, 37, 8);

        // Set data for each chart (optional)
        sheet.Charts[chartIndex1].NSeries.Add("B2:B4", true);
        sheet.Charts[chartIndex1].NSeries.CategoryData = "A2:A4";

        sheet.Charts[chartIndex2].NSeries.Add("B2:B4", true);
        sheet.Charts[chartIndex2].NSeries.CategoryData = "A2:A4";

        sheet.Charts[chartIndex3].NSeries.Add("B2:B4", true);
        sheet.Charts[chartIndex3].NSeries.CategoryData = "A2:A4";

        // Remove the third chart (zero‑based index 2)
        sheet.Charts.RemoveAt(2);

        // Save the workbook
        workbook.Save("RemovedThirdChart.xlsx");
    }
}