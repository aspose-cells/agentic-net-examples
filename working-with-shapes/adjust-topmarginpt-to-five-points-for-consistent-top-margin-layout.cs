// Title: Aspose.Cells for .NET – Set Shape Text Top Margin to 5 pt
// Description: Creates a new workbook, adds a rectangle shape, assigns text, sets the shape's TextBody.TextAlignment.TopMarginPt to 5 points for uniform spacing, and saves the file as TopMarginAdjusted.xlsx.
// Keywords: Aspose.Cells .NET shape margin | TopMarginPt C# example | set shape text top margin Aspose.Cells | Excel rectangle shape formatting | C# Aspose.Cells TextAlignment | adjust shape text spacing | global Excel automation | US developers Aspose.Cells | European .NET Excel library | GitHub Aspose.Cells sample
// Common Searches: How to set shape text top margin in Aspose.Cells C# | Aspose.Cells TopMarginPt property example | C# code to adjust rectangle shape text margin in Excel | Set 5 point top margin for shape text using Aspose.Cells | Aspose.Cells shape formatting tutorial | Excel shape text spacing Aspose.Cells .NET
// Developer Intent: Set the top margin of a shape’s text to exactly 5 points using Aspose.Cells for .NET.
// Use Cases: Designing Excel reports where shape captions require consistent vertical padding. | Automating bulk formatting of multiple shapes across worksheets with a fixed top margin. | Creating printable dashboards where text inside shapes must align uniformly. | Generating localized Excel templates that need precise text positioning inside shapes.
// AI Prompts: Provide a C# snippet that sets TextBody.TextAlignment.TopMarginPt to 5 for a specific shape in Aspose.Cells. | Show how to loop through all shapes in a worksheet and apply a 5‑point top margin to their text. | Explain the difference between TopMarginPt, BottomMarginPt, and internal padding in Aspose.Cells TextAlignment. | Give guidance on troubleshooting missing top margin changes when saving an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a new workbook, adds a rectangle shape, assigns text, sets the shape's TextBody.TextAlignment.TopMarginPt to 5 points for uniform spacing, and saves the file as TopMarginAdjusted.xlsx.
    public class AdjustTopMarginDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, upper left offset in pixels,
                // height in pixels, width in pixels, rotation angle
                Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 100, 200, 0);

                // Set sample text for the shape
                shape.Text = "Sample Text with Top Margin";

                // Adjust the top margin of the shape's text to 5 points
                shape.TextBody.TextAlignment.TopMarginPt = 5.0;

                // Save the workbook
                string outputPath = "TopMarginAdjusted.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AdjustTopMarginDemo.Run();
        }
    }
}
