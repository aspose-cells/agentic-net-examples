// Title: Add a Rectangle Shape with Text to an Excel Worksheet using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new Workbook, insert a rectangle shape on the first worksheet, set custom text, apply a solid light‑green fill and a 1.5‑pt solid border, and save the file as an .xlsx document.
// Keywords: Aspose.Cells rectangle shape C# | add shape with text Aspose.Cells | Excel shape fill color .NET | draw rectangle on worksheet canvas | Aspose.Cells shape styling
// Common Searches: how to add a rectangle shape with text in Aspose.Cells | Aspose.Cells set fill color and border for a shape | C# code to insert labeled rectangles in Excel | Aspose.Cells shape formatting examples | add custom shape to worksheet using Aspose.Cells
// Developer Intent: Insert a rectangle shape onto a worksheet, assign custom text, and customize its fill and border properties.
// Use Cases: Highlight key sections in a generated report with colored, labeled boxes. | Create a simple dashboard layout by placing labeled containers for charts or tables. | Add instructional callout boxes to guide end‑users within an Excel file.
// AI Prompts: Show how to add a rectangle shape with multiline text that auto‑sizes to fit the content using Aspose.Cells for .NET. | Provide C# code to generate multiple labeled rectangles in a loop, align them horizontally, and apply consistent styling.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a new Workbook, insert a rectangle shape on the first worksheet, set custom text, apply a solid light‑green fill and a 1.5‑pt solid border, and save the file as an .xlsx document.
    public class AddRectangleWithTextDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet canvas
                // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
                RectangleShape rectangle = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 120, 200);

                // Set custom text for the rectangle shape
                rectangle.Text = "This is a custom rectangle shape";

                // Customize appearance (fill color, line style)
                rectangle.Fill.FillType = FillType.Solid;
                rectangle.Fill.SolidFill.Color = System.Drawing.Color.LightGreen;
                rectangle.Line.DashStyle = MsoLineDashStyle.Solid;
                rectangle.Line.Weight = 1.5;

                // Define output file path
                string outputPath = "RectangleWithTextDemo.xlsx";

                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook to a file
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
            AddRectangleWithTextDemo.Run();
        }
    }
}
