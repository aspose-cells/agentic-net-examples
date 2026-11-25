// Title: Hide data labels for the third series in an Aspose.Cells column chart using C#
// Common Searches: C# Aspose.Cells hide data labels for a specific series in a column chart | How to remove labels from the third series of an Excel chart with Aspose.Cells | Aspose.Cells chart NSeries disable data labels for one series | Set DataLabels.IsDeleted true for a series in Aspose.Cells C# example | Control individual series data labels in Aspose.Cells column chart
// Tags: Aspose.Cells series data label suppression C# | Chart.NSeries.DataLabels.IsDeleted usage | disable third series labels Aspose.Cells | column chart label control Aspose.Cells | C# Excel chart series formatting Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace DisableDataLabelsExample
{
    // The example creates a workbook with three data series, adds a column chart, enables data labels for the first two series, disables data labels for the third series by setting DataLabels.IsDeleted to true, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for three series
            // Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // Series 1
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Series 2
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Series 3
            sheet.Cells["D1"].PutValue("Series 3");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the three series to the chart
            // Series 1
            chart.NSeries.Add("B2:B4", true);
            // Series 2
            chart.NSeries.Add("C2:C4", true);
            // Series 3
            chart.NSeries.Add("D2:D4", true);

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A4";

            // Disable data labels for the third series (index 2)
            // Two ways: hide all labels or delete the DataLabels object.
            // Here we set IsDeleted = true which removes the labels completely.
            chart.NSeries[2].DataLabels.IsDeleted = true;

            // Optionally, enable data labels for the first two series for demonstration
            chart.NSeries[0].DataLabels.ShowValue = true;
            chart.NSeries[1].DataLabels.ShowValue = true;

            // Save the workbook
            string outputPath = "ChartWithThirdSeriesLabelsDisabled.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
