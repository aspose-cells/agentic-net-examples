// Title: Apply a two‑color horizontal linear gradient to an Excel chart background with Aspose.Cells for .NET (C#)
// Common Searches: c# aspose.cells set chart background to horizontal two‑color gradient | how to use FillFormat.SetTwoColorGradient for Excel chart area in Aspose.Cells | example of applying dark blue and orange gradient to chart area using Aspose.Cells | Aspose.Cells chart area gradient fill tutorial for .NET developers | create column chart with gradient background in C# using Aspose.Cells | set chart area FillFormat gradient type to horizontal in Aspose.Cells
// Tags: Aspose.Cells chart area two‑color gradient | C# FillFormat SetTwoColorGradient method | horizontal linear gradient fill Excel chart | chart background gradient Aspose.Cells .NET | Excel chart FillFormat gradient configuration | Aspose.Cells linear gradient style chart

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartGradientBackgroundDemo
{
    // The example creates a workbook, adds a column chart with sample data, and configures the chart area's background to a horizontal two‑color linear gradient (DarkBlue to Orange) using the FillFormat API, then saves the file as ChartGradientBackground.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("Alpha");
                worksheet.Cells["A3"].PutValue("Beta");
                worksheet.Cells["A4"].PutValue("Gamma");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["B3"].PutValue(55);
                worksheet.Cells["B4"].PutValue(80);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure the chart title (optional)
                chart.Title.Text = "Sample Column Chart";

                // Access the chart area (the whole chart background)
                // ChartArea is a ChartFrame; its Area property provides FillFormat
                ChartArea chartArea = chart.ChartArea;
                Area area = chartArea.Area;
                FillFormat fill = area.FillFormat;

                // Ensure the fill type is set to Gradient
                fill.FillType = FillType.Gradient;

                // Apply a two‑color linear gradient (horizontal) with contrasting colors
                // Example colors: DarkBlue and Orange
                Color colorStart = Color.DarkBlue;
                Color colorEnd = Color.Orange;
                GradientStyleType gradientStyle = GradientStyleType.Horizontal; // Linear horizontal gradient
                int variant = 1; // Variant 1 corresponds to the first preset in the Fill Effects dialog

                // Set the gradient fill
                fill.SetTwoColorGradient(colorStart, colorEnd, gradientStyle, variant);

                // Save the workbook to a file
                string outputPath = "ChartGradientBackground.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart with gradient background:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
