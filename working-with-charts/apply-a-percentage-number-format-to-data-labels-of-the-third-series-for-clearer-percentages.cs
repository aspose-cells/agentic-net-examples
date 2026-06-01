using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (categories + three series)
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        // Series 1 values
        worksheet.Cells["B1"].PutValue("Series1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Series 2 values
        worksheet.Cells["C1"].PutValue("Series2");
        worksheet.Cells["C2"].PutValue(15);
        worksheet.Cells["C3"].PutValue(25);
        worksheet.Cells["C4"].PutValue(35);

        // Series 3 values (as fractions to be shown as percentages)
        worksheet.Cells["D1"].PutValue("Series3");
        worksheet.Cells["D2"].PutValue(0.1);
        worksheet.Cells["D3"].PutValue(0.2);
        worksheet.Cells["D4"].PutValue(0.3);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = worksheet.Charts[chartIndex];

        // Add three series to the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries[0].Name = "Series1";

        chart.NSeries.Add("C2:C4", true);
        chart.NSeries[1].Name = "Series2";

        chart.NSeries.Add("D2:D4", true);
        chart.NSeries[2].Name = "Series3";

        // Set category (X‑axis) data
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the third series and apply a percentage format
        Series thirdSeries = chart.NSeries[2];
        thirdSeries.DataLabels.ShowValue = true;          // show the value
        thirdSeries.DataLabels.ShowPercentage = true;    // optional: show percentage flag
        thirdSeries.DataLabels.NumberFormat = "0.00%";   // percentage number format

        // Save the workbook
        workbook.Save("ThirdSeriesPercentageDataLabels.xlsx");
    }
}