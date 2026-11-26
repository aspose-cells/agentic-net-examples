// Title: Bring the rectangle shape named "ChartOverlay" to the front layer in an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# move a specific shape to the top of the z-order | How to set a shape as frontmost in an Excel worksheet with Aspose.Cells .NET | C# Aspose.Cells bring overlay rectangle above chart objects | Using ToFrontOrBack method to prioritize a named shape in Aspose.Cells | Aspose.Cells example for ordering shapes in an Excel workbook
// Tags: Aspose.Cells shape z-order manipulation | C# Aspose.Cells bring shape to front | Aspose.Cells ToFrontOrBack method usage | Excel rectangle overlay positioning Aspose.Cells | Aspose.Cells chart overlay shape front layer

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ShapeFrontDemo
{
    // The example creates a workbook, adds a column chart, inserts a rectangle shape named 'ChartOverlay' with a semi‑transparent yellow fill, and calls ToFrontOrBack(1000) to move the shape to the topmost layer before saving the file as ShapeFrontDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some data for a chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Sample Chart";

                // Add a rectangle shape that will serve as the overlay
                // Parameters: topRow, top, leftColumn, left, height, width
                Shape overlayShape = sheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 150);
                overlayShape.Name = "ChartOverlay";

                // Give the shape a visible fill (semi‑transparent yellow) and hide its border
                overlayShape.Fill.FillType = FillType.Solid;
                overlayShape.Fill.SolidFill.Color = System.Drawing.Color.FromArgb(128, System.Drawing.Color.Yellow);
                overlayShape.HasLine = false; // Hide border

                // Bring the shape to the front layer so it appears above all other objects
                // Positive value moves the shape forward; a large number guarantees topmost position
                overlayShape.ToFrontOrBack(1000);

                // Save the workbook
                workbook.Save("ShapeFrontDemo.xlsx");
                Console.WriteLine("Workbook saved. Shape 'ChartOverlay' is now on the front layer.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
