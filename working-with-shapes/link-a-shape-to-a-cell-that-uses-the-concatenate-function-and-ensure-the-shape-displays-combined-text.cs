using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedToConcatenate
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put sample data in B1 and C1
            sheet.Cells["B1"].PutValue("Hello, ");
            sheet.Cells["C1"].PutValue("World!");

            // Set A1 formula to concatenate B1 and C1
            sheet.Cells["A1"].Formula = "=CONCATENATE(B1,C1)";

            // Add a rectangle shape (using MsoDrawingType as ShapeType may be unavailable)
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);

            // Link the shape to cell A1 (the cell containing the CONCATENATE formula)
            // The two boolean parameters indicate whether to update the shape when the cell changes
            shape.SetLinkedCell("A1", true, true);

            // Force the shape to refresh its displayed value based on the linked cell
            shape.UpdateSelectedValue();

            // Define output file path
            string outputPath = "ShapeLinkedToConcatenate.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}