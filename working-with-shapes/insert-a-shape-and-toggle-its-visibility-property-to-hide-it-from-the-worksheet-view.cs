// Title: Insert a rectangle shape and hide it in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add rectangle shape and hide it in Excel | how to use shape.IsHidden property with Aspose.Cells workbook | hide a drawing shape from worksheet view using Aspose.Cells .NET | Aspose.Cells API to make an Excel shape invisible in C# | set shape visibility to hidden when creating Excel file with Aspose.Cells
// Tags: add rectangle shape Aspose.Cells C# | shape.IsHidden property Aspose.Cells | hide drawing shape Excel Aspose.Cells | Aspose.Cells worksheet shape visibility | C# Excel shape invisibility Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a rectangle shape to the first worksheet, assigns text to the shape, sets its IsHidden property to true to hide it from the worksheet view, and saves the file as output.xlsx.
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

            // Add a rectangle shape at row 1, column 0 with size 100x50 points
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Shape type
                1,   // Upper left row
                0,   // Upper left column
                0,   // Top offset (points)
                0,   // Left offset (points)
                100, // Height (points)
                50   // Width (points)
            );

            // Set text for the shape
            shape.Text = "Hidden Shape";

            // Hide the shape from the worksheet view
            shape.IsHidden = true;

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
