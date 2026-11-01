// Title: How to programmatically change a worksheet shape's Z-order to front and back using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a rectangle shape to a worksheet, moves it to the front layer, then moves it to the back layer, and prints the ZOrderPosition after each operation. | Show how to read and set the ZOrderPosition property of a Shape object in Aspose.Cells to control shape layering in an Excel file.
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to bring shape to front | How to send a shape to the back layer in an Excel workbook using Aspose.Cells .NET | Retrieve initial Z-order of a shape after adding it with Aspose.Cells | Example of changing shape layering order in Aspose.Cells C#
// Tags: Aspose.Cells shape ZOrderPosition API | C# worksheet shape layering with Aspose.Cells | modify shape Z-order in Excel using Aspose.Cells | Aspose.Cells shape front/back ordering | track shape Z-order values Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, records its initial Z‑order, moves the shape to the front by setting ZOrderPosition to the highest index, then moves it to the back by setting ZOrderPosition to zero, prints each Z‑order value, and saves the workbook as ShapeZOrderDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet and obtain the Shape object
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // Record initial Z-order value
            int initialZOrder = shape.ZOrderPosition;
            Console.WriteLine($"Initial Z-order: {initialZOrder}");

            // Bring the shape to the front by setting its Z-order to the highest position
            shape.ZOrderPosition = sheet.Shapes.Count - 1;
            int frontZOrder = shape.ZOrderPosition;
            Console.WriteLine($"Z-order after bringing to front: {frontZOrder}");

            // Send the shape to the back by setting its Z-order to the lowest position
            shape.ZOrderPosition = 0;
            int backZOrder = shape.ZOrderPosition;
            Console.WriteLine($"Z-order after sending to back: {backZOrder}");

            // Save the workbook
            string outputPath = "ShapeZOrderDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
