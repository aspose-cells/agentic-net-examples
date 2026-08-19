// Title: C# – Export Aspose.Cells Chart to Landscape PDF (Full‑Width)
// Description: Creates a workbook, adds sample data, inserts a column chart, sets the chart's PageSetup orientation to Landscape, and exports the chart to a PDF so it occupies the entire page width.
// Keywords: Aspose.Cells chart PDF landscape | C# export chart to PDF | set chart orientation Aspose.Cells | chart.ToPdf landscape | Aspose.Cells PageSetup orientation
// Common Searches: Aspose.Cells set chart to landscape before PDF export | C# export column chart as landscape PDF | How to use PageSetup orientation with Aspose.Cells chart | chart.ToPdf landscape mode example
// Developer Intent: Configure a chart's page orientation to Landscape and generate a PDF file using Aspose.Cells for .NET.
// Use Cases: Produce printable reports where wide charts need the full page width. | Create dashboard PDFs that display column charts without clipping. | Automate batch conversion of multiple charts into separate landscape PDFs.
// AI Prompts: Write C# code with Aspose.Cells that sets a chart's orientation to Landscape and saves it as a PDF. | Explain the steps to modify a chart's PageSetup before calling ToPdf in Aspose.Cells. | Show how to iterate through all charts in a workbook and export each to a landscape‑oriented PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart, sets the chart's PageSetup orientation to Landscape, and exports the chart to a PDF so it occupies the entire page width.
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
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set the chart's page orientation to Landscape
                chart.PageSetup.Orientation = PageOrientationType.Landscape;

                // Export the chart to a PDF file; the landscape orientation ensures full width usage
                chart.ToPdf("ChartLandscape.pdf");

                Console.WriteLine("Chart exported to PDF with landscape orientation successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ChartToPdfLandscapeDemo.Run();
        }
    }
}
