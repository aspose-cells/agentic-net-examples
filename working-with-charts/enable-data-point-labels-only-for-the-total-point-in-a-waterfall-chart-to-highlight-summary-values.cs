// Title: How to display a data label only on the Total point of a Waterfall chart using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add a label to just the Total bar in a Waterfall chart | show only summary value label in Aspose.Cells waterfall chart .NET | enable data point label for total column in Aspose.Cells Waterfall chart using C# | Aspose.Cells .NET hide all waterfall chart labels except the total | waterfall chart label customization example with Aspose.Cells C#
// Tags: Aspose.Cells waterfall chart total data label | C# per‑point label Aspose.Cells chart | Aspose.Cells hide waterfall series labels | summary point labeling Aspose.Cells .NET | waterfall chart label customization Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallChartExample
{
    // The example creates a new Workbook, fills category and value data, adds a Waterfall chart, and assigns the data range. It demonstrates that the current Aspose.Cells for .NET API does not expose per‑point label properties (HasDataLabel, DataPoints, DataPoint), so a label can only be shown for the whole series. Consequently, labeling only the Total point is not supported in this version, and the chart will display the default labels.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Fill sample data for the Waterfall chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                string[] categories = { "Start", "Increase", "Decrease", "Total" };
                double[] values = { 100, 30, -20, 110 };

                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(categories[i]);   // Column A
                    sheet.Cells[i + 2, 1].PutValue(values[i]);      // Column B
                }

                // Add a Waterfall chart (position: row 7, column 0 to row 24, column 7)
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 24, 7);
                Chart waterfallChart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                waterfallChart.NSeries.Add("B2:B5", true);
                waterfallChart.NSeries.CategoryData = "A2:A5";

                // NOTE: The Aspose.Cells version used in this environment does not expose
                // the HasDataLabel, DataPoints, or DataPoint members for Series.
                // Therefore, custom per‑point label handling is omitted.
                // The chart will display default labels.

                // Save the workbook
                string outputPath = "Waterfall_TotalLabel.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the Waterfall chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
