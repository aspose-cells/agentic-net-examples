// Title: Create a Waterfall chart in Aspose.Cells for .NET, hide default data label text, and set up a summary range for custom labels
// Common Searches: how to create a waterfall chart with Aspose.Cells in C# and hide the default data labels | Aspose.Cells C# example for suppressing data label values in a waterfall chart | using Aspose.Cells to prepare a summary range for custom waterfall chart labels | C# code to generate a waterfall chart and customize data label visibility with Aspose.Cells | Aspose.Cells .NET tutorial for creating a waterfall chart and adjusting label placeholders
// Tags: Aspose.Cells waterfall chart generation | C# hide waterfall chart data labels | Aspose.Cells custom label summary range | Aspose.Cells data label visibility settings | Aspose.Cells save workbook with chart

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new Workbook, populates columns A‑C with category, value, and label data, adds a Waterfall chart using the values range B2:B5, disables the display of values, category names, and series names for the chart’s data labels, and saves the workbook as WaterfallChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Populate data for the waterfall chart
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("Label"); // summary range for data labels

            string[] categories = { "Start", "Increase", "Decrease", "End" };
            double[] values = { 100, 30, -20, 110 };
            string[] labels = { "Start", "Add 30", "Subtract 20", "Total 110" };

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(categories[i]);   // Column A
                sheet.Cells[i + 2, 1].PutValue(values[i]);      // Column B
                sheet.Cells[i + 2, 2].PutValue(labels[i]);      // Column C (summary range)
            }

            // -------------------------------------------------
            // Add a Waterfall chart
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the series (values only)
            chart.NSeries.Add("B2:B5", true);

            // -------------------------------------------------
            // Configure data labels (hide default content)
            // -------------------------------------------------
            chart.NSeries[0].DataLabels.ShowValue = false;
            chart.NSeries[0].DataLabels.ShowCategoryName = false;
            chart.NSeries[0].DataLabels.ShowSeriesName = false;

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            string outputPath = "WaterfallChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
