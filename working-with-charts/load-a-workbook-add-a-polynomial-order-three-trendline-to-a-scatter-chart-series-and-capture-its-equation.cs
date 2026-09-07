// Title: Add a cubic (order‑3) polynomial series to a scatter chart and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, generates a column of X³ values, adds this column as a new series to the first scatter chart, and saves the workbook to a new file. | Modify the program to calculate the cubic polynomial equation for the added series and output the equation to the console.
// Common Searches: how to add a third order polynomial series to a scatter chart with Aspose.Cells C# | Aspose.Cells create cubic trendline for scatter chart in .NET | C# Aspose.Cells add custom series to existing Excel chart and save workbook | retrieve polynomial equation from Aspose.Cells chart series .NET
// Tags: Aspose.Cells add cubic series to scatter chart | C# create polynomial series in Excel chart | Aspose.Cells write X³ values to worksheet | save modified workbook with Aspose.Cells .NET | Excel scatter chart custom series Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example ensures an input.xlsx file exists (creating one with X values and a quadratic Y series if needed), loads the workbook, computes X³ values in a new column, adds those values as a cubic (order‑3) series to the first scatter chart, and saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a simple workbook with sample data and a scatter chart.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Name = "Data";

                // Populate sample data for X and Y series.
                ws.Cells["A1"].PutValue("X");
                ws.Cells["B1"].PutValue("Y");
                for (int i = 0; i < 10; i++)
                {
                    ws.Cells[i + 1, 0].PutValue(i);                     // X values
                    ws.Cells[i + 1, 1].PutValue(Math.Pow(i, 2));       // Y = X^2 (quadratic)
                }

                // Add a scatter chart.
                int chartIdx = ws.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
                Chart scatterChart = ws.Charts[chartIdx];
                scatterChart.NSeries.Add("B2:B11", true);
                scatterChart.NSeries[0].XValues = "A2:A11";

                wb.Save(inputPath);
            }

            // Load the workbook safely.
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook '{inputPath}': {ex.Message}");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Assume the first chart is the target scatter chart.
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            Chart targetChart = worksheet.Charts[0];

            // Ensure the chart has at least one series.
            if (targetChart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Add a polynomial (order 3) approximation by creating a new series.
            try
            {
                // Prepare column for cubic values.
                worksheet.Cells["C1"].PutValue("Y³");
                for (int row = 2; row <= 11; row++)
                {
                    double x = worksheet.Cells[row - 1, 0].DoubleValue; // X from column A
                    worksheet.Cells[row - 1, 2].PutValue(Math.Pow(x, 3)); // Y³ in column C
                }

                // Add the new series to the chart.
                int newSeriesIdx = targetChart.NSeries.Add("C2:C11", true);
                Series polySeries = targetChart.NSeries[newSeriesIdx];
                polySeries.XValues = "A2:A11";
                polySeries.Name = "Cubic Approximation";

                // Optional styling can be added here if supported by the API version.

                Console.WriteLine("Polynomial (order 3) series added to the chart.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to add polynomial series: " + ex.Message);
            }

            // Ensure output directory exists.
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
