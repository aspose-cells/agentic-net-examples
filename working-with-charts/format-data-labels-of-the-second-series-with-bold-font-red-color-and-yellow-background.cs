using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for two series
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add the two series to the chart
        chart.NSeries.Add("B2:B4", true); // First series
        chart.NSeries.Add("C2:C4", true); // Second series
        chart.NSeries.CategoryData = "A2:A4";

        // Access the second series (index 1)
        Series secondSeries = chart.NSeries[1];

        // Enable data labels for the second series
        secondSeries.DataLabels.ShowValue = true;

        // Format the data labels: bold font, red color
        secondSeries.DataLabels.Font.IsBold = true;
        secondSeries.DataLabels.Font.Color = Color.Red;

        // Set yellow background for the data labels
        secondSeries.DataLabels.Area.BackgroundColor = Color.Yellow;

        // Apply the font settings to all child nodes (optional but ensures consistency)
        secondSeries.DataLabels.ApplyFont();

        // Save the workbook
        workbook.Save("FormattedDataLabels.xlsx", SaveFormat.Xlsx);
    }
}