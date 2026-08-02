// Title: Set Shape Text Left Margin to 8 Points with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a rectangle shape, assign text, and set the shape's left text margin to 8 points using ShapeTextAlignment.LeftMarginPt in C#. The result is saved as an .xlsx file.
// Keywords: Aspose.Cells | C# | ShapeTextAlignment | LeftMarginPt | shape text margin | set left margin points | rectangle shape | Excel shape padding | align text with cell borders | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set left margin of shape text | ShapeTextAlignment LeftMarginPt C# | add padding to shape text in Excel using Aspose | set shape text left padding Aspose.Cells .NET | align shape label with cell border Aspose.Cells
// Developer Intent: Apply an 8‑point left margin to a shape’s text so it aligns with the cell border.
// Use Cases: Create report cards where shape labels start exactly at the cell’s left edge. | Standardize 8‑point left padding for callout shapes in a financial dashboard. | Generate invoices with shapes whose text respects a consistent left margin.
// AI Prompts: Write C# code that loops through all shapes in a worksheet and sets ShapeTextAlignment.LeftMarginPt to 8 points using Aspose.Cells. | Explain the relationship between points, millimeters, and inches for ShapeTextAlignment margins and show conversion formulas in Aspose.Cells. | Provide a sample that reads a left‑margin value from appsettings.json and applies it to each shape’s text alignment in an Excel file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a rectangle shape, assign text, and set the shape's left text margin to 8 points using ShapeTextAlignment.LeftMarginPt in C#. The result is saved as an .xlsx file.
    public class SetLeftMarginPtDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape that will contain text
                // Parameters: upper left row, upper left column, top, left, width, height
                Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 200, 100);
                shape.Text = "Text aligned with cell borders";

                // Access the text alignment of the shape and set the left margin to 8 points
                ShapeTextAlignment alignment = shape.TextBody.TextAlignment;
                alignment.LeftMarginPt = 8.0; // 8 points

                // Define output file path
                string outputPath = "SetLeftMarginPtDemo.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetLeftMarginPtDemo.Run();
        }
    }
}
