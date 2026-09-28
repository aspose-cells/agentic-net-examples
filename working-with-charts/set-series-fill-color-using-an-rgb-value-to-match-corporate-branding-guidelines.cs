// Title: Set a specific RGB fill color for a chart series in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a column chart with Aspose.Cells, adds sample data, and assigns the first series a solid fill using Color.FromArgb(79,129,189). | Show how to access a chart series' Area.FillFormat.SolidFill.Color property in Aspose.Cells to apply a corporate branding RGB color. | Demonstrate saving the workbook after customizing the series fill color so the color appears in the generated XLSX file.
// Common Searches: how to change the fill color of a chart series in Aspose.Cells C# | Aspose.Cells set series solid fill to specific RGB value | C# Aspose.Cells column chart custom series color example | apply corporate branding color to Excel chart series using Aspose.Cells .NET
// Tags: Aspose.Cells chart series solid fill color | C# set series fill RGB Aspose.Cells | Excel column chart custom series color .NET | Area.FillFormat.SolidFill.Color usage Aspose.Cells | Save workbook with colored chart series Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesFillColorDemo
{
    // The example creates a new workbook, populates it with sample data, adds a column chart, and sets the first series' fill color to the RGB value (79,129,189) via the Area.FillFormat.SolidFill.Color property, then saves the file as SeriesFillColorDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set the fill color of the first series using an RGB value (corporate branding color)
                Color corporateColor = Color.FromArgb(79, 129, 189);
                Series series = chart.NSeries[0];

                // Apply fill color to the series
                series.Area.FillFormat.SolidFill.Color = corporateColor;

                // Save the workbook
                string outputPath = "SeriesFillColorDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
