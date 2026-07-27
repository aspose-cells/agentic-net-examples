// Title: Export Aspose.Cells Chart to PDF in Landscape Orientation (C#)
// Description: This example creates a workbook, adds a column chart with sample data, sets the chart's page orientation to Landscape, and exports the chart to a PDF. The landscape setting ensures the chart uses the full width of the page for better readability and printing.
// Keywords: Aspose.Cells chart to PDF | C# landscape PDF export | chart PageSetup orientation | Aspose.Cells ToPdf landscape | export chart full width PDF | .NET chart PDF orientation | Aspose.Cells PDF page setup
// Common Searches: Aspose.Cells export chart to PDF landscape C# | set chart orientation to landscape before PDF export | how to use chart.ToPdf with landscape page in .NET | Aspose.Cells chart full width PDF | C# Aspose.Cells PDF page orientation for charts
// Developer Intent: Set a chart’s orientation to Landscape and generate a PDF file.
// Use Cases: Create printable reports where charts span the entire width of a landscape page. | Generate dashboard PDFs that require horizontal layout for better data visualization. | Produce PDFs for presentations that match landscape printing standards.
// AI Prompts: Show C# code that sets a chart’s PageSetup orientation to Landscape and exports it to PDF using Aspose.Cells. | Explain how to configure Aspose.Cells chart orientation before calling ToPdf for a full‑width landscape PDF. | Provide a step‑by‑step guide to export an Aspose.Cells chart to a landscape PDF in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a column chart with sample data, sets the chart's page orientation to Landscape, and exports the chart to a PDF. The landscape setting ensures the chart uses the full width of the page for better readability and printing.
    public class ChartToPdfLandscapeDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Fruits");
            worksheet.Cells["A3"].PutValue("Vegetables");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(50);
            worksheet.Cells["B3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Set the chart's page orientation to Landscape
            chart.PageSetup.Orientation = PageOrientationType.Landscape;

            // Define output PDF path
            string outputPath = "ChartLandscape.pdf";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the chart to a PDF file; the landscape orientation ensures full width usage
            chart.ToPdf(outputPath);

            Console.WriteLine("Chart exported to PDF with landscape orientation successfully.");
        }
    }
}
