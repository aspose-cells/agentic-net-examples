// Title: C# – Apply Radial Gradient Fill with Three Color Stops to an Aspose.Cells Column Chart
// Description: Shows how to create a workbook, add sample data, insert a column chart, and set the first series to a radial gradient fill with three color stops (red, yellow, blue) using Aspose.Cells for .NET, then save the file as RadialGradientChart.xlsx.
// Keywords: Aspose.Cells | C# | radial gradient fill | chart gradient stops | column chart styling | Excel gradient fill API | GradientFillType.Radial | Aspose.Cells chart formatting | gradient fill Aspose.Cells .NET
// Common Searches: Aspose.Cells radial gradient fill C# | how to add gradient stops to a chart series in Aspose.Cells | column chart gradient fill Aspose.Cells .NET example | set radial gradient for Excel chart using Aspose | Aspose.Cells chart series fill type gradient
// Developer Intent: Create a column chart and apply a radial gradient fill that transitions through three colors to its data series.
// Use Cases: Design a sales performance workbook where column heights are highlighted with a red‑yellow‑blue radial gradient for visual impact. | Build an executive dashboard Excel file that uses brand‑specific radial gradients on charts to differentiate key metrics. | Improve the aesthetic of exported Excel reports by replacing solid fills with smooth radial gradients on chart series.
// AI Prompts: Generate code that changes the three gradient stops to semi‑transparent colors while keeping the radial gradient on a column chart series. | Provide a C# example that applies a radial gradient with five color stops to a line chart using Aspose.Cells. | Explain how to define a reusable radial gradient style and apply it to multiple series across different worksheets in a workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add sample data, insert a column chart, and set the first series to a radial gradient fill with three color stops (red, yellow, blue) using Aspose.Cells for .NET, then save the file as RadialGradientChart.xlsx.
public class RadialGradientChartDemo
{
    public static void Run()
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
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(20);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series and set its fill type to gradient
            Series series = chart.NSeries[0];
            series.Area.FillFormat.FillType = FillType.Gradient;

            // Configure the gradient as radial
            GradientFill gradientFill = series.Area.FillFormat.GradientFill;
            gradientFill.SetGradient(GradientFillType.Radial, 0, GradientDirectionType.FromCenter);

            // Clear any existing gradient stops
            gradientFill.GradientStops.Clear();

            // Add three color stops for a smooth transition
            // Position is expressed as a percentage (0.0 to 1.0)
            gradientFill.GradientStops.Add(0.0, Color.Red, 255);      // Start with opaque red
            gradientFill.GradientStops.Add(0.5, Color.Yellow, 255);   // Middle transition to yellow
            gradientFill.GradientStops.Add(1.0, Color.Blue, 255);    // End with opaque blue

            // Save the workbook
            string outputPath = "RadialGradientChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        RadialGradientChartDemo.Run();
    }
}
