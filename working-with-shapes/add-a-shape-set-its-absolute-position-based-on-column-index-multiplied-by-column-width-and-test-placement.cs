// Title: Add a rectangle shape to an Excel worksheet and position it using column index multiplied by column width in Aspose.Cells for .NET
// AI Prompts: Generate C# code that inserts a rectangle shape on the first worksheet, obtains the pixel width of a specified column, computes left = columnIndex * columnWidth, assigns the value to Shape.Left and Shape.Top, and saves the workbook. | Write a reusable method that accepts a column index and shape dimensions, uses Worksheet.Cells.GetColumnWidthPixel to calculate the absolute horizontal coordinate, sets the shape's position, and returns the configured Shape object. | Create a console application that prints the calculated horizontal coordinate and the column's pixel width, then verifies the placement by saving the file as an .xlsx workbook.
// Common Searches: Aspose.Cells C# set shape horizontal position using column width in pixels | How to calculate absolute shape coordinates from column index in Aspose.Cells | Retrieve column width in pixels and place a rectangle shape at that column with Aspose.Cells .NET
// Tags: Aspose.Cells shape left positioning | Aspose.Cells column width in pixels | Aspose.Cells worksheet shape placement | Aspose.Cells calculate absolute shape coordinates | Aspose.Cells C# shape positioning example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, reads the pixel width of a given column, adds a rectangle shape, sets its Left property to columnIndex multiplied by the column width, outputs the computed values for verification, and saves the file as ShapePositionTest.xlsx.
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

            // Define the target column index (0‑based)
            int columnIndex = 2; // Example: third column

            // Retrieve the width of the column in pixels
            double columnWidthPx = sheet.Cells.GetColumnWidthPixel(columnIndex);

            // Calculate the absolute left position: column index * column width
            double leftPosition = columnIndex * columnWidthPx;

            // Add a rectangle shape (width: 80 points, height: 40 points)
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 0, 0, 0, 0, 80, 40);

            // Set the shape's absolute position (cast to int because Shape.Left expects int)
            shape.Left = (int)leftPosition;   // Horizontal position
            shape.Top = 10;                    // Small top offset for visibility

            // Output the calculated values for verification
            Console.WriteLine($"Shape Left = {shape.Left} (expected {(int)leftPosition})");
            Console.WriteLine($"Column {columnIndex} width (pixels) = {columnWidthPx}");

            // Save the workbook to verify the shape placement
            string outputPath = "ShapePositionTest.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
