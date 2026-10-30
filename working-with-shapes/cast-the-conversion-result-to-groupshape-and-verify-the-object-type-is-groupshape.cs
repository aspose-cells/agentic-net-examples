// Title: How to cast a worksheet Shape to GroupShape and verify its type using Aspose.Cells for .NET (C#)
// Common Searches: C# cast worksheet shape to GroupShape with Aspose.Cells | check if a shape is a GroupShape in Aspose.Cells .NET | determine shape type at runtime using Aspose.Cells API | using the 'as' operator to convert Shape to GroupShape in C# | verify GroupShape after adding an AutoShape to a workbook
// Tags: shape to groupshape casting Aspose.Cells | runtime shape type checking Aspose.Cells | groupshape detection in worksheet C# | aspnet cells shape hierarchy verification | as operator shape conversion Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a rectangle AutoShape to the first worksheet, attempts to cast the Shape to a GroupShape using the C# 'as' operator, and then checks whether the cast succeeded by testing for null, outputting the result accordingly.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle AutoShape to the worksheet.
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = sheet.Shapes.AddAutoShape(
                AutoShapeType.Rectangle, // shape type
                0,   // upper left row
                0,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                200, // height (pixels)
                200  // width (pixels)
            );

            // Attempt to cast the shape to GroupShape
            GroupShape? groupShape = shape as GroupShape;

            // Verify whether the cast succeeded
            if (groupShape != null)
            {
                Console.WriteLine("The object is successfully cast to GroupShape.");
            }
            else
            {
                Console.WriteLine("The object is NOT a GroupShape.");
            }

            // (Optional) Save the workbook if needed – follows lifecycle rule
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
