// Title: Convert a Column Chart to a Line Chart and Export as PDF with Aspose.Cells for .NET
// Description: Shows how to build a workbook, add a column chart, programmatically switch its type to a line chart, and save the chart alone as a PDF file using the Aspose.Cells .NET API.
// Keywords: Aspose.Cells | C# chart manipulation | change chart type | column to line chart | export chart to PDF | Chart.ToPdf | .NET workbook | chart API | single chart PDF export | Aspose.Cells example
// Common Searches: Aspose.Cells change chart type C# | export single chart to PDF Aspose.Cells | convert column chart to line chart programmatically | C# Aspose.Cells ToPdf example | how to switch chart type before PDF export
// Developer Intent: Modify an existing column chart to a line chart and generate a PDF that contains only the transformed chart.
// Use Cases: Create a line‑chart report from raw worksheet data without saving the full workbook. | Allow users to select a preferred chart style and instantly export the chosen chart as a PDF. | Automate chart type conversion in batch processes and produce standalone PDF visualizations.
// AI Prompts: Write C# code with Aspose.Cells that changes a chart's type to Line and saves it as a PDF. | Explain how to export only a specific chart to PDF while keeping its formatting intact. | Show how to use a variable to set the chart type dynamically and then call ToPdf.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to build a workbook, add a column chart, programmatically switch its type to a line chart, and save the chart alone as a PDF file using the Aspose.Cells .NET API.
class ChangeChartTypeAndExportPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Fruits");
        worksheet.Cells["A3"].PutValue("Vegetables");
        worksheet.Cells["A4"].PutValue("Grains");

        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(50);
        worksheet.Cells["B3"].PutValue(30);
        worksheet.Cells["B4"].PutValue(20);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Change the chart type from Column to Line
        chart.Type = ChartType.Line;

        // Export the chart to a PDF file
        chart.ToPdf("LineChart.pdf");

        Console.WriteLine("Chart has been changed to Line type and exported to PDF.");
    }
}
