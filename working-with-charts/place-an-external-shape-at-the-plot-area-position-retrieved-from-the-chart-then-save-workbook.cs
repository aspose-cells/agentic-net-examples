// Title: Overlay a rectangle shape on a chart’s plot area and save the workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells place rectangle shape exactly over chart plot area | how to get chart plot area coordinates with Aspose.Cells and add a shape | Aspose.Cells add external shape to Excel chart and save workbook | retrieve plot area X Y width height from Aspose.Cells chart in C# | save Excel file after inserting shape over chart using Aspose.Cells .NET
// Tags: Aspose.Cells add rectangle shape to chart plot area | C# retrieve chart PlotArea dimensions Aspose.Cells | overlay external shape on Excel chart using Aspose.Cells | save workbook after shape insertion Aspose.Cells .NET | MsoDrawingType.Rectangle placement with chart coordinates | Excel chart plot area coordinates C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a column chart, reads the chart's PlotArea X, Y, Width, and Height values, inserts a rectangle shape positioned precisely over that plot area using the retrieved coordinates, and then saves the workbook to an Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a sample column chart to have a plot area
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Retrieve plot area position and size (X,Y are left/top)
                float plotLeft = chart.PlotArea.X;
                float plotTop = chart.PlotArea.Y;
                float plotWidth = chart.PlotArea.Width;
                float plotHeight = chart.PlotArea.Height;

                // Place a rectangle shape exactly over the plot area
                // ShapeCollection.AddShape expects row/column indices and offsets in points.
                // Use row=0, column=0 and cast float values to int.
                Shape externalShape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    0,                     // upper left row
                    0,                     // upper left column
                    (int)plotTop,          // top offset
                    (int)plotLeft,         // left offset
                    (int)plotHeight,       // height
                    (int)plotWidth);       // width

                // Optional: add some text to identify the shape
                externalShape.Text = "External Shape";

                // Ensure output directory exists
                string outputPath = "output.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? ".";
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
