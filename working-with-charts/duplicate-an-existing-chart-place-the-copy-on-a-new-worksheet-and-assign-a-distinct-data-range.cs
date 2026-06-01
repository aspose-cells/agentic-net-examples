using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet (source sheet)
        Workbook workbook = new Workbook();
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Populate source data (A1:B4)
        sourceSheet.Cells["A1"].PutValue("Category");
        sourceSheet.Cells["B1"].PutValue("Value");
        sourceSheet.Cells["A2"].PutValue("A");
        sourceSheet.Cells["B2"].PutValue(10);
        sourceSheet.Cells["A3"].PutValue("B");
        sourceSheet.Cells["B3"].PutValue(20);
        sourceSheet.Cells["A4"].PutValue("C");
        sourceSheet.Cells["B4"].PutValue(30);

        // Add a chart on the source sheet
        int srcChartIndex = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart srcChart = sourceSheet.Charts[srcChartIndex];
        srcChart.SetChartDataRange("A1:B4", true);
        srcChart.Title.Text = "Source Chart";

        // Add a new worksheet that will hold the duplicated chart
        Worksheet destSheet = workbook.Worksheets.Add("Copy");

        // Populate distinct data for the copied chart (C1:D4)
        destSheet.Cells["C1"].PutValue("Category");
        destSheet.Cells["D1"].PutValue("Value");
        destSheet.Cells["C2"].PutValue("X");
        destSheet.Cells["D2"].PutValue(40);
        destSheet.Cells["C3"].PutValue("Y");
        destSheet.Cells["D3"].PutValue(50);
        destSheet.Cells["C4"].PutValue("Z");
        destSheet.Cells["D4"].PutValue(60);

        // Add a chart on the destination sheet with the same type and position as the source chart
        int destChartIndex = destSheet.Charts.Add(srcChart.Type, 5, 0, 15, 5);
        Chart destChart = destSheet.Charts[destChartIndex];

        // Assign the distinct data range to the duplicated chart
        destChart.SetChartDataRange("C1:D4", true);
        destChart.Title.Text = "Copied Chart";

        // Save the workbook
        workbook.Save("ChartDuplicate.xlsx");
    }
}