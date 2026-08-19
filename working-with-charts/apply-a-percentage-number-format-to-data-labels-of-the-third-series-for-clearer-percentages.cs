// Title: Apply Percentage Number Format to Data Labels of the Third Series in an Aspose.Cells Column Chart (C#)
// Description: Creates a workbook with three series, adds a column chart, enables data labels for the third series, and applies the number format "0.00%" so the labels display as percentages before saving the file.
// Keywords: Aspose.Cells | C# | .NET | percentage number format | chart data labels | third series | column chart | NumberFormat property | Excel chart formatting | Aspose.Cells for .NET
// Common Searches: Aspose.Cells set data label number format | C# format chart series as percentage | show percentage labels for specific series Aspose.Cells | column chart third series percentage Aspose.Cells | NumberFormat property chart data labels .NET
// Developer Intent: Display the values of the third series as percentages by applying a percentage number format to its data labels.
// Use Cases: Generate a multi‑series column chart where only the ratio series needs percentage labels. | Create Excel reports that highlight proportion data while keeping other series in default format. | Export charts from .NET applications with customized label formatting for clearer visual analysis.
// AI Prompts: Write C# code using Aspose.Cells to set the NumberFormat of the third series data labels to "0.00%". | Show how to enable ShowValue and ShowPercentage for a specific chart series without affecting other series. | Explain how to change the decimal precision of percentage labels for a chosen series in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook with three series, adds a column chart, enables data labels for the third series, and applies the number format "0.00%" so the labels display as percentages before saving the file.
class ApplyPercentageFormatToThirdSeries
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for three series
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        // Series 1 values
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Series 2 values
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Series 3 values (we will format its data labels as percentages)
        sheet.Cells["D1"].PutValue("Series3");
        sheet.Cells["D2"].PutValue(0.1);   // 10%
        sheet.Cells["D3"].PutValue(0.2);   // 20%
        sheet.Cells["D4"].PutValue(0.3);   // 30%

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add the three series to the chart
        chart.NSeries.Add("B2:B4", true); // Series 1
        chart.NSeries.Add("C2:C4", true); // Series 2
        chart.NSeries.Add("D2:D4", true); // Series 3

        // Set category (X) data
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the third series
        Series thirdSeries = chart.NSeries[2]; // zero‑based index
        thirdSeries.DataLabels.ShowValue = true;          // show the value
        thirdSeries.DataLabels.ShowPercentage = true;    // optional: show percentage text
        thirdSeries.DataLabels.NumberFormat = "0.00%";   // apply percentage number format

        // Save the workbook
        workbook.Save("ChartWithThirdSeriesPercentageLabels.xlsx");
    }
}
