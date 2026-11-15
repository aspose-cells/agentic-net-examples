// Title: Save an Excel worksheet containing a column chart as an SVG file using Aspose.Cells for .NET
// Common Searches: c# aspnet save excel worksheet chart to svg with aspose.cells | how to export column chart from workbook to svg using Aspose.Cells .NET | convert Excel chart to scalable vector graphic programmatically in C# | Aspose.Cells SaveFormat.Svg example for chart export | generate SVG image of Excel column chart in a .NET application
// Tags: Aspose.Cells SaveFormat.Svg usage | C# Excel chart SVG generation | Aspose.Cells worksheet SVG export | column chart to SVG with Aspose.Cells | SVG output for Excel charts .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills cells A1:B5 with category and value data, adds a column chart linked to that range, sets a chart title, and saves the entire worksheet—including the chart—as an SVG file by using Workbook.Save with SaveFormat.Svg.
class ChartToSvgExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            using (Workbook workbook = new Workbook())
            {
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                // A1:A5 - Categories, B1:B5 - Values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(130);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 27, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart series
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Optional: set chart title
                chart.Title.Text = "Monthly Sales";

                // Save the entire worksheet (including the chart) as SVG
                string svgPath = "Worksheet.svg";
                workbook.Save(svgPath, SaveFormat.Svg);
                Console.WriteLine($"Worksheet saved as SVG: {Path.GetFullPath(svgPath)}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
