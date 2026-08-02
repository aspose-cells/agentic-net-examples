// Title: Set a rectangle shape's top text margin to 5 points with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a rectangle shape, assign text, and set the shape's TextBody.TextAlignment.TopMarginPt property to 5 points before saving the file.
// Keywords: Aspose.Cells | .NET | C# | shape text margin | TopMarginPt | rectangle shape | Excel workbook | TextBody | TextAlignment | margin points
// Common Searches: Aspose.Cells set shape top margin C# | TopMarginPt property example | adjust rectangle text margin points Aspose.Cells | how to change top margin of shape text in Excel using .NET | C# Aspose.Cells shape text alignment
// Developer Intent: Apply a 5‑point top margin to the text inside a rectangle shape.
// Use Cases: Design a report template where all shape labels start exactly 5 points below the shape edge for consistent visual spacing. | Programmatically enforce uniform top margins across multiple shapes in a generated Excel dashboard. | Create a printable worksheet where text inside shapes aligns with corporate layout guidelines.
// AI Prompts: Generate C# code that sets TopMarginPt for every shape in an existing workbook to a user‑defined value. | Explain how TopMarginPt differs from other margin properties in Aspose.Cells shape text formatting. | Show a step‑by‑step tutorial for adjusting shape text margins using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a rectangle shape, assign text, and set the shape's TextBody.TextAlignment.TopMarginPt property to 5 points before saving the file.
    public class AdjustTopMarginPtDemo
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
                shape.Text = "Sample text with top margin";

                // Adjust the top margin of the shape's text to 5 points
                shape.TextBody.TextAlignment.TopMarginPt = 5.0;

                // Save the workbook to a file
                string outputPath = "AdjustTopMarginPtDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            AdjustTopMarginPtDemo.Run();
        }
    }
}
