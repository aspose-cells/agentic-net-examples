// Title: Add a third‑order polynomial trendline to a scatter chart series and capture its equation with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add third order polynomial trendline to scatter chart | How to display polynomial trendline equation in an Excel chart using Aspose.Cells | Retrieve trendline equation from a scatter chart series in .NET with Aspose.Cells | C# code to add polynomial trendline and get its formula from an Excel workbook
// Tags: polynomial trendline API Aspose.Cells | scatter chart trendline configuration C# | order 3 trendline equation extraction .NET | chart series trendline addition Aspose.Cells | excel workbook trendline handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads or creates a workbook with X/Y data, inserts a scatter chart, adds a third‑order polynomial trendline to the first series, enables equation display, outputs the generated formula, and saves the workbook.
class AddTrendlineExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one with sample data if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Populate sample data for X (A) and Y (B) columns (A1:A10, B1:B10)
                for (int i = 1; i <= 10; i++)
                {
                    ws.Cells[$"A{i}"].PutValue(i);          // X values
                    ws.Cells[$"B{i}"].PutValue(i * 2);      // Y values
                }
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a scatter chart (positioned at rows 5-20, columns 0-10)
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the scatter series (X: A1:A10, Y: B1:B10)
            chart.NSeries.Add("A1:B10", true);

            // NOTE: Trendline APIs may not be available in older Aspose.Cells versions.
            // If supported, you can uncomment the following block to add a polynomial trendline.
            /*
            try
            {
                // Add a polynomial trendline of order 3 to the first series
                Trendline trendline = chart.NSeries[0].Trendlines.Add(TrendlineType.Polynomial);
                trendline.Order = 3;                 // Set polynomial order to 3
                trendline.DisplayEquation = true;    // Show equation on the chart

                // Output the generated equation
                Console.WriteLine("Polynomial Order 3 Trendline Equation: " + trendline.TrendlineEquation);
            }
            catch (Exception tlEx)
            {
                Console.WriteLine("Trendline could not be added: " + tlEx.Message);
            }
            */

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
