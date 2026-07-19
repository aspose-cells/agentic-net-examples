// Title: Aspose.Cells C# – Convert a Column Chart to a Line Chart and Export as PDF
// Description: Demonstrates how to create a workbook, add a column chart, programmatically switch its type to a line chart using the Chart.Type property, and save the resulting chart directly to a PDF file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart type conversion | C# line chart PDF export | change column chart to line | Aspose.Cells ToPdf method | .NET chart manipulation | dynamic chart style Aspose.Cells
// Common Searches: Aspose.Cells change chart type C# | Export line chart to PDF with Aspose.Cells | How to convert column chart to line chart programmatically | Chart.Type property Aspose.Cells example | Save Aspose.Cells chart as PDF
// Developer Intent: Switch an existing column chart to a line chart in code and generate a PDF containing the updated chart.
// Use Cases: Create PDF reports where the visual style of charts can be altered at runtime. | Provide end‑users with a UI to select chart styles and instantly export the chosen format. | Automate batch conversion of multiple column charts to line charts for standardized documentation.
// AI Prompts: Generate C# code that changes a column chart to a line chart with Aspose.Cells and saves it as PDF. | Explain the steps to modify Chart.Type and use ToPdf for exporting a chart in Aspose.Cells .NET. | Write a reusable method that accepts a worksheet and chart index, converts the chart to a line type, and returns the PDF file path.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTypeChange
{
    // Demonstrates how to create a workbook, add a column chart, programmatically switch its type to a line chart using the Chart.Type property, and save the resulting chart directly to a PDF file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Fruits");
            sheet.Cells["A3"].PutValue("Vegetables");
            sheet.Cells["A4"].PutValue("Grains");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(20);

            // Add a column chart (initial type)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Change the chart type from Column to Line
            chart.Type = ChartType.Line;

            // Export the chart to a PDF file
            chart.ToPdf("ChartLine.pdf");

            Console.WriteLine("Chart type changed to Line and exported to PDF successfully.");
        }
    }
}
