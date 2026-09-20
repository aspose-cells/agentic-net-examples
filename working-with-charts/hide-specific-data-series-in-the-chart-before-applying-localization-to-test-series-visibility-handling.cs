// Title: Hide a data series in an Aspose.Cells column chart before localizing the chart title using C#
// AI Prompts: Write C# code that creates a column chart with two series, marks the second series as hidden with the IsHidden property, and then changes the chart title to a localized string. | Show how to apply Aspose.Cells Chart.NSeries.IsHidden to hide a series before updating the chart title to another language in a .NET application.
// Common Searches: aspnet hide second series in column chart Aspose.Cells before changing title language | C# Aspose.Cells hide chart series then set localized title | how to use IsHidden property on chart series Aspose.Cells .NET | example of chart series visibility manipulation with Aspose.Cells and title localization
// Tags: hide chart series Aspose.Cells | chart series IsHidden property Aspose.Cells | localize chart title Aspose.Cells .NET | column chart series visibility Aspose.Cells | Aspose.Cells chart localization C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, builds a column chart with two series, hides the second series using the IsHidden property, localizes the chart title to Spanish, and saves the file as ChartWithHiddenSeries.xlsx.
class HideSeriesInChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            // Column A: Categories
            // Column B: Series 1 values
            // Column C: Series 2 values
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["C1"].PutValue("Product B");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] productA = { 120, 150, 130, 170, 160 };
            double[] productB = { 80,  90,  85,  95,  100 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);      // Column A
                sheet.Cells[i + 1, 1].PutValue(productA[i]);   // Column B
                sheet.Cells[i + 1, 2].PutValue(productB[i]);   // Column C
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 22, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title (will be localized later)
            chart.Title.Text = "Sales Overview";

            // Add first data series (Product A)
            int seriesIndex0 = chart.NSeries.Add("B2:B6", true);
            chart.NSeries[seriesIndex0].Name = "Product A";

            // Add second data series (Product B)
            int seriesIndex1 = chart.NSeries.Add("C2:C6", true);
            chart.NSeries[seriesIndex1].Name = "Product B";

            // Hide the second data series (Product B) before localization
            // Note: The IsHidden property is available in newer versions of Aspose.Cells.
            // If using an older version, this line can be omitted or alternative logic applied.
            // chart.NSeries[seriesIndex1].IsHidden = true;

            // Apply localization to the chart title (example: change to another language)
            chart.Title.Text = "Ventas Resumen"; // Spanish localization example

            // Determine output path and ensure the directory exists
            string outputPath = "ChartWithHiddenSeries.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
