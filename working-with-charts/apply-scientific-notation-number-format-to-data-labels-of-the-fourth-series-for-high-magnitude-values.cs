using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ApplyScientificNotationToFourthSeries
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for four series (columns B to E) with high magnitude values
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        // Series 1
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(1_200_000);
        sheet.Cells["B3"].PutValue(1_500_000);
        sheet.Cells["B4"].PutValue(1_800_000);

        // Series 2
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(2_200_000);
        sheet.Cells["C3"].PutValue(2_500_000);
        sheet.Cells["C4"].PutValue(2_800_000);

        // Series 3
        sheet.Cells["D1"].PutValue("Series3");
        sheet.Cells["D2"].PutValue(3_200_000);
        sheet.Cells["D3"].PutValue(3_500_000);
        sheet.Cells["D4"].PutValue(3_800_000);

        // Series 4 (the target series)
        sheet.Cells["E1"].PutValue("Series4");
        sheet.Cells["E2"].PutValue(4_200_000);
        sheet.Cells["E3"].PutValue(4_500_000);
        sheet.Cells["E4"].PutValue(4_800_000);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add the four series to the chart
        chart.NSeries.Add("B2:B4", true); // Series 1
        chart.NSeries.Add("C2:C4", true); // Series 2
        chart.NSeries.Add("D2:D4", true); // Series 3
        chart.NSeries.Add("E2:E4", true); // Series 4

        // Set category (X) data
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the fourth series and apply scientific notation format
        Series fourthSeries = chart.NSeries[3]; // zero‑based index, 3 = fourth series
        fourthSeries.DataLabels.ShowValue = true;
        fourthSeries.DataLabels.NumberFormat = "0.00E+00"; // scientific notation

        // Save the workbook
        workbook.Save("ChartWithScientificNotation.xlsx");
    }
}