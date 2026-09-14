// Title: Convert an Aspose.Cells column chart to a PDF with landscape page orientation using C#
// AI Prompts: Generate C# code that creates a worksheet, adds a column chart, sets Chart.PageSetup.Orientation to Landscape, and saves the chart as a PDF with Aspose.Cells. | Show how to export an Aspose.Cells chart to PDF in landscape mode so the chart occupies the full page width.
// Common Searches: c# aspnet export excel chart to pdf landscape orientation using aspose.cells | how to set chart page orientation to landscape before pdf conversion in asp.net | asp.net core generate column chart and save as landscape pdf with aspose.cells | aspose.cells chart to pdf full width landscape example c#
// Tags: Aspose.Cells chart to PDF landscape | C# set chart orientation Aspose.Cells | Export Excel chart as landscape PDF .NET | Chart.PageSetup.Orientation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, inserts sample data, builds a column chart, changes the chart's PageSetup.Orientation to Landscape, and then exports the chart to a PDF file using Aspose.Cells for .NET.
    public class ChartToPdfLandscapeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.SetChartDataRange("A1:B4", true);
                chart.Title.Text = "Sample Chart";

                // Set the chart's page orientation to Landscape to use full width
                chart.PageSetup.Orientation = PageOrientationType.Landscape;

                // Define output PDF path
                string outputPath = "ChartLandscape.pdf";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export the chart to a PDF file
                chart.ToPdf(outputPath);

                Console.WriteLine("Chart exported to PDF with landscape orientation.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Application entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
