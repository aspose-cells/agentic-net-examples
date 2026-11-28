// Title: Add exponential trendline to a chart series in an XLS workbook and extract its equation using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add exponential trendline to existing chart and get equation text | retrieve trendline equation from an XLS chart with Aspose.Cells .NET API | display exponential trendline equation on a column chart in Excel 97-2003 workbook using Aspose.Cells | force chart calculation after adding a trendline in Aspose.Cells C# | create a chart if missing and add exponential trendline with Aspose.Cells example
// Tags: Aspose.Cells exponential trendline C# | chart series trendline equation extraction Aspose.Cells | add trendline to XLS chart Aspose.Cells | retrieve trendline data label text Aspose.Cells | force chart calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AddExponentialTrendline
{
    // The sample loads an existing XLS file or creates a new workbook, ensures a column chart exists, adds an exponential trendline to the first series, configures the trendline to display its equation, forces chart calculation, reads the equation text from the trendline's data label, and saves the workbook in Excel 97‑2003 format using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files
            string inputPath = "input.xls";
            string outputPath = "output.xls";

            try
            {
                Workbook workbook;

                // Load existing workbook if the file exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    Console.WriteLine($"Input file '{inputPath}' not found. A new workbook will be created.");
                }

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one chart; if not, create a simple column chart
                Chart chart;
                if (worksheet.Charts.Count > 0)
                {
                    chart = worksheet.Charts[0];
                }
                else
                {
                    // Populate sample data
                    worksheet.Cells["A1"].PutValue("Category");
                    worksheet.Cells["B1"].PutValue("Value");
                    for (int i = 2; i <= 5; i++)
                    {
                        worksheet.Cells[$"A{i}"].PutValue($"Item{i - 1}");
                        worksheet.Cells[$"B{i}"].PutValue(i * 10);
                    }

                    // Add a column chart
                    int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                    chart = worksheet.Charts[chartIndex];
                    chart.NSeries.Add("B2:B5", true);
                    chart.NSeries.CategoryData = "A2:A5";
                }

                // Verify that the chart contains at least one series
                if (chart.NSeries.Count == 0)
                {
                    Console.WriteLine("The chart does not contain any series to which a trendline can be added.");
                    return;
                }

                Series series = chart.NSeries[0];

                // Add an exponential trendline to the series
                int trendlineIndex = series.TrendLines.Add(TrendlineType.Exponential);
                Trendline trendline = series.TrendLines[trendlineIndex];

                // Configure the trendline to display its equation
                trendline.DisplayEquation = true;
                trendline.DisplayRSquared = false; // optional: hide R²

                // Force chart calculation so that data labels (including the equation) are generated
                chart.Calculate();

                // Retrieve the equation text from the trendline's data label
                string equationText = trendline.DataLabels.Text;

                Console.WriteLine("Exponential Trendline Equation:");
                Console.WriteLine(equationText ?? "Equation text not available.");

                // Save the modified workbook in Excel 97-2003 format
                workbook.Save(outputPath, SaveFormat.Excel97To2003);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
