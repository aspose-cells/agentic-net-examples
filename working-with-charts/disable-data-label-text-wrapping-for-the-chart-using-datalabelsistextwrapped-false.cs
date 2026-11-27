// Title: Disable text wrapping for chart data labels using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# disable data label text wrap in column chart | set DataLabels.IsTextWrapped false Aspose.Cells example | prevent chart data label wrapping with Aspose.Cells .NET | how to turn off text wrapping for series labels in Aspose.Cells chart | Aspose.Cells chart label formatting without wrapping in C#
// Tags: Aspose.Cells chart data label no wrap | DataLabels.IsTextWrapped property .NET | disable chart label wrapping C# | column chart data label formatting Aspose.Cells | Aspose.Cells chart customization XLSX

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace DisableDataLabelWrappingDemo
{
    // This example creates a workbook, adds a column chart, enables data labels, sets DataLabels.IsTextWrapped to false to stop label text from wrapping, applies a long custom label, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the resulting workbook will be saved
            const string outputPath = "ChartDataLabels_NoWrap.xlsx";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(15000);
            sheet.Cells["B3"].PutValue(20000);
            sheet.Cells["B4"].PutValue(18000);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;               // Show the numeric value
            dataLabels.Position = LabelPositionType.OutsideEnd;

            // Disable text wrapping for the data labels
            dataLabels.IsTextWrapped = false;

            // Optional: set a longer custom text to demonstrate that wrapping is disabled
            dataLabels.Text = "VeryLongCategoryNameThatWouldNormallyWrapIfEnabled";

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'. Data label text wrapping is disabled.");
        }
    }
}
