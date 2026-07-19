// Title: Export Aspose.Cells Chart to PDF in Landscape Orientation (C#)
// Description: Creates a workbook, adds a column chart, sets the chart's PageSetup.Orientation to Landscape, and exports the chart to a PDF so it uses the full page width.
// Keywords: Aspose.Cells chart PDF landscape | C# Aspose.Cells export chart to PDF | set chart orientation Aspose.Cells | .NET chart to PDF landscape | chart.PageSetup.Orientation | chart.ToPdf Aspose.Cells
// Common Searches: Aspose.Cells set chart to landscape before PDF export | C# export chart as PDF landscape Aspose.Cells | how to change chart orientation to landscape in Aspose.Cells | export column chart to PDF full width Aspose.Cells .NET | chart.ToPdf landscape mode example
// Developer Intent: Set a chart’s orientation to landscape and generate a PDF file with the chart occupying the full page width.
// Use Cases: Produce printable PDF reports where charts need a wider layout. | Create dashboard documentation with landscape‑oriented chart PDFs. | Automate batch conversion of workbook charts to individual landscape PDFs for archiving.
// AI Prompts: Generate C# code that builds a pie chart with Aspose.Cells, sets its orientation to landscape, and saves it as a PDF. | Explain step‑by‑step how to change a chart’s PageSetup orientation to landscape before calling ToPdf in Aspose.Cells for .NET. | Show how to export multiple charts from a workbook to separate landscape‑oriented PDF files using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart, sets the chart's PageSetup.Orientation to Landscape, and exports the chart to a PDF so it uses the full page width.
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
