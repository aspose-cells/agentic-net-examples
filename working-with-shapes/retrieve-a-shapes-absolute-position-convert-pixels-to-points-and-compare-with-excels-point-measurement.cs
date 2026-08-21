// Title: Get Shape Position in Pixels, Convert to Points, and Compare with Excel Values using Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds a rectangle shape, reads its X/Y coordinates and size in pixels, converts those values to points using the 96 DPI → 72 points factor, and compares the results with the WidthPt and HeightPt properties reported by Excel. The differences are printed and the workbook is saved.
// Keywords: Aspose.Cells | C# | shape position pixels | pixel to point conversion | WidthPt | HeightPt | DPI 96 | Excel shape dimensions | retrieve shape coordinates | compare shape size | worksheet shapes
// Common Searches: Aspose.Cells get shape X Y in pixels | convert shape pixel size to points C# | compare WidthPt with pixel conversion Aspose.Cells | shape absolute position Aspose.Cells .NET | pixel to point factor 96 DPI Aspose
// Developer Intent: Obtain a shape’s pixel coordinates and dimensions, translate them to points, and verify the conversion against Excel’s point measurements.
// Use Cases: Validate that programmatically added shapes align with Excel’s layout by matching pixel‑derived points to WidthPt/HeightPt. | Generate exact point measurements for shapes when preparing documents for printing or PDF export. | Debug positioning mismatches when converting legacy pixel‑based layouts to point‑based formats.
// AI Prompts: Write C# code with Aspose.Cells that reads a shape’s X, Y, Width, and Height in pixels and converts them to points using a 96 DPI to 72 points conversion factor. | Explain why Aspose.Cells’s WidthPt and HeightPt values may differ slightly from pixel‑to‑point calculations. | Suggest how to adjust a shape’s position programmatically so that the converted point values match a target Excel layout.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds a rectangle shape, reads its X/Y coordinates and size in pixels, converts those values to points using the 96 DPI → 72 points factor, and compares the results with the WidthPt and HeightPt properties reported by Excel. The differences are printed and the workbook is saved.
    public class ShapePositionComparison
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape at row 2, column 2 (zero‑based indexes) with no offset,
                // height 100 pixels and width 150 pixels
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 150);

                // Position in pixels (offset from worksheet's top‑left corner)
                int shapeXPixel = shape.X;
                int shapeYPixel = shape.Y;

                // Size reported by Excel in points
                double shapeWidthPt = shape.WidthPt;
                double shapeHeightPt = shape.HeightPt;

                // Conversion factor from pixels to points (96 DPI → 72 points per inch)
                const double dpi = 96.0;
                const double pointsPerInch = 72.0;
                double pixelsToPointsFactor = pointsPerInch / dpi; // 0.75

                // Convert pixel values to points
                double shapeXPt = shapeXPixel * pixelsToPointsFactor;
                double shapeYPt = shapeYPixel * pixelsToPointsFactor;
                double shapeWidthPx = shape.Width;   // width in pixels
                double shapeHeightPx = shape.Height; // height in pixels

                double shapeWidthPtFromPx = shapeWidthPx * pixelsToPointsFactor;
                double shapeHeightPtFromPx = shapeHeightPx * pixelsToPointsFactor;

                // Output comparison
                Console.WriteLine("Shape Position and Size Comparison:");
                Console.WriteLine($"Position X: {shapeXPixel} px  => {shapeXPt:F2} pt (converted)");
                Console.WriteLine($"Position Y: {shapeYPixel} px  => {shapeYPt:F2} pt (converted)");
                Console.WriteLine($"Width:       {shapeWidthPx} px => {shapeWidthPtFromPx:F2} pt (converted)");
                Console.WriteLine($"Height:      {shapeHeightPx} px => {shapeHeightPtFromPx:F2} pt (converted)");
                Console.WriteLine();
                Console.WriteLine($"Excel reported WidthPt:  {shapeWidthPt:F2} pt");
                Console.WriteLine($"Excel reported HeightPt: {shapeHeightPt:F2} pt");
                Console.WriteLine();
                Console.WriteLine("Difference (converted - Excel):");
                Console.WriteLine($"Width difference:  {Math.Abs(shapeWidthPtFromPx - shapeWidthPt):F2} pt");
                Console.WriteLine($"Height difference: {Math.Abs(shapeHeightPtFromPx - shapeHeightPt):F2} pt");

                // Save the workbook
                string outputPath = "ShapePositionComparison.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShapePositionComparison.Run();
        }
    }
}
