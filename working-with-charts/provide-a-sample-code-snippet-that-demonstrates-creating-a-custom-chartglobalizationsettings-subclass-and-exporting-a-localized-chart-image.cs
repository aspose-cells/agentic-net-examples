// Title: C# Aspose.Cells example: create a French‑localized column chart and save it as a transparent PNG image
// AI Prompts: Write C# code that builds a workbook with month and sales data, adds a column chart, sets French titles for the chart, axes, and legend, and exports the chart to a PNG file using Aspose.Cells. | Show how to configure ImageOrPrintOptions for a transparent background and one page per sheet when rendering an Aspose.Cells chart to an image. | Demonstrate setting the legend position to the bottom and customizing axis titles before calling Chart.ToImage in Aspose.Cells.
// Common Searches: Aspose.Cells C# export chart with French axis titles to PNG | how to save an Aspose.Cells chart as a transparent PNG image | C# example of localized chart titles using Aspose.Cells | set legend position bottom Aspose.Cells chart image export
// Tags: export chart to PNG Aspose.Cells C# | set French chart titles Aspose.Cells | transparent background ImageOrPrintOptions | legend position bottom Aspose.Cells chart | column chart localization Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartLocalization
{
    // The sample creates a workbook, fills it with month and sales data, adds a column chart, applies French titles to the chart, category axis, value axis, and legend, positions the legend at the bottom, and exports the chart as a transparent PNG image using ImageOrPrintOptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                using (Workbook workbook = new Workbook())
                {
                    // Access the first worksheet
                    Worksheet sheet = workbook.Worksheets[0];

                    // Populate sample data
                    sheet.Cells["A1"].PutValue("Mois");
                    sheet.Cells["B1"].PutValue("Ventes");

                    string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
                    double[] sales = { 1200, 1500, 1800, 1300, 1700, 1600 };

                    for (int i = 0; i < months.Length; i++)
                    {
                        sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                        sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
                    }

                    // Add a column chart
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 10);
                    Chart chart = sheet.Charts[chartIndex];

                    // Set data source
                    chart.NSeries.Add("B2:B7", true);
                    chart.NSeries.CategoryData = "A2:A7";

                    // Set localized titles
                    chart.Title.Text = "Ventes Mensuelles";          // French for "Monthly Sales"
                    chart.CategoryAxis.Title.Text = "Mois";          // French for "Month"
                    chart.ValueAxis.Title.Text = "Valeur de Vente"; // French for "Sales Value"
                    chart.Legend.Position = LegendPositionType.Bottom; // Position the legend at the bottom
                    chart.Legend.Text = "Produit A";

                    // Export the chart as an image
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        OnePagePerSheet = true,
                        Transparent = true
                    };

                    // Directly save the chart image to a file
                    chart.ToImage("LocalizedChart.png", imgOptions);

                    Console.WriteLine("Localized chart image has been saved as 'LocalizedChart.png'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
