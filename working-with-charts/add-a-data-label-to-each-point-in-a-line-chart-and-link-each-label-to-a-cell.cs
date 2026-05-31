using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AddDataLabelsLinkedToCells
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["C1"].PutValue("Label");

        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["A5"].PutValue("Apr");

        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);
        sheet.Cells["B5"].PutValue(200);

        // Custom label text for each point (will be linked)
        sheet.Cells["C2"].PutValue("Q1");
        sheet.Cells["C3"].PutValue("Q2");
        sheet.Cells["C4"].PutValue("Q3");
        sheet.Cells["C5"].PutValue("Q4");

        // Add a line chart
        int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set data range for the series
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Enable data labels for the series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = false;               // hide default value
        series.DataLabels.ShowCellRange = true;            // use cell range as label text
        series.DataLabels.LinkedSource = "C2:C5";          // link each label to corresponding cell
        series.DataLabels.Position = LabelPositionType.Above; // position suitable for line chart
        series.DataLabels.NumberFormatLinked = true;       // keep number format in sync if needed

        // Optional: calculate the chart to ensure labels are rendered
        chart.Calculate();

        // Save the workbook
        workbook.Save("LineChartWithLinkedDataLabels.xlsx");
    }
}