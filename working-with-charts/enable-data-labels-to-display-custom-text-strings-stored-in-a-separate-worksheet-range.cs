using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class CustomDataLabelDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Worksheet 1 – source data for the chart
        // -------------------------------------------------
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Category (X) values
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["A2"].PutValue("Alpha");
        dataSheet.Cells["A3"].PutValue("Beta");
        dataSheet.Cells["A4"].PutValue("Gamma");

        // Numeric (Y) values
        dataSheet.Cells["B1"].PutValue("Value");
        dataSheet.Cells["B2"].PutValue(120);
        dataSheet.Cells["B3"].PutValue(85);
        dataSheet.Cells["B4"].PutValue(150);

        // -------------------------------------------------
        // Worksheet 2 – custom text strings for data labels
        // -------------------------------------------------
        Worksheet labelSheet = workbook.Worksheets[workbook.Worksheets.Add()];
        labelSheet.Name = "Labels";

        // Custom label strings (must correspond to each data point)
        labelSheet.Cells["A1"].PutValue("Custom Labels");
        labelSheet.Cells["A2"].PutValue("High");
        labelSheet.Cells["A3"].PutValue("Medium");
        labelSheet.Cells["A4"].PutValue("Low");

        // -------------------------------------------------
        // Add a column chart to the data sheet
        // -------------------------------------------------
        int chartIndex = dataSheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
        Chart chart = dataSheet.Charts[chartIndex];

        // Bind the series to the numeric values
        chart.NSeries.Add("B2:B4", true);
        // Bind the category (X) axis
        chart.NSeries.CategoryData = "A2:A4";

        // -------------------------------------------------
        // Configure data labels to use the custom strings
        // -------------------------------------------------
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = false;          // hide the default numeric value
        series.DataLabels.ShowCellRange = true;       // enable cell range linking
        // Reference the range that contains custom label texts
        series.DataLabels.LinkedSource = "Labels!A2:A4";

        // Optional: format appearance of the data labels
        series.DataLabels.Font.Color = System.Drawing.Color.Blue;
        series.DataLabels.Position = LabelPositionType.InsideEnd;

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("CustomDataLabels.xlsx");
    }
}