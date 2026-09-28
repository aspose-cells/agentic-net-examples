// Title: How to display fourth series data labels in scientific notation on a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Set the NumberFormat of the fourth series' data labels to "0.00E+00" in an Aspose.Cells column chart. | Enable data labels only for the last series and apply an exponential format to those labels in C#. | Create a multi‑series column chart and format series 4 labels so they appear as 1.23E+06.
// Common Searches: Aspose.Cells C# column chart format data labels of a specific series as scientific notation | How to apply a custom number format to only one series in an Aspose.Cells chart | Show large numbers in chart labels using exponential format with Aspose.Cells .NET | C# example for setting scientific notation on fourth series data labels in an Excel workbook
// Tags: column chart data labels exponential format Aspose.Cells | set series-specific number format .NET | fourth series label formatting Excel Aspose | exponential number format chart labels C# | high‑value chart series formatting Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsScientificNotationDemo
{
    // The sample creates a workbook with four data series, adds a column chart, enables data labels only for the fourth series, sets its NumberFormat to "0.00E+00" so the labels render in scientific (exponential) notation, and saves the file as ScientificNotationDataLabels.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with high magnitude values for four series
            // Categories (X‑axis)
            sheet.Cells["A2"].PutValue("Category 1");
            sheet.Cells["A3"].PutValue("Category 2");
            sheet.Cells["A4"].PutValue("Category 3");

            // Series 1
            sheet.Cells["B2"].PutValue(1_200_000);
            sheet.Cells["B3"].PutValue(2_500_000);
            sheet.Cells["B4"].PutValue(3_800_000);

            // Series 2
            sheet.Cells["C2"].PutValue(4_100_000);
            sheet.Cells["C3"].PutValue(5_600_000);
            sheet.Cells["C4"].PutValue(6_900_000);

            // Series 3
            sheet.Cells["D2"].PutValue(7_200_000);
            sheet.Cells["D3"].PutValue(8_500_000);
            sheet.Cells["D4"].PutValue(9_800_000);

            // Series 4 (the one we will format)
            sheet.Cells["E2"].PutValue(10_100_000);
            sheet.Cells["E3"].PutValue(11_400_000);
            sheet.Cells["E4"].PutValue(12_700_000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the four series to the chart
            // Series are added by specifying the values range; the category range is shared
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries.Add("D2:D4", true);
            chart.NSeries.Add("E2:E4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the fourth series only
            Series fourthSeries = chart.NSeries[3]; // zero‑based index
            fourthSeries.DataLabels.ShowValue = true;

            // Apply scientific notation format to the data labels of the fourth series
            // Example format: 1.23E+06
            fourthSeries.DataLabels.NumberFormat = "0.00E+00";

            // Save the workbook
            workbook.Save("ScientificNotationDataLabels.xlsx");
        }
    }
}
