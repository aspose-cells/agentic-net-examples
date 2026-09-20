// Title: Lock rectangle shapes in an Aspose.Cells worksheet and reposition a locked shape programmatically (C#)
// AI Prompts: Generate C# code that creates an Excel workbook with Aspose.Cells, adds two rectangle shapes, sets IsLocked = true for each shape, moves the first shape by changing its UpperLeftRow and UpperLeftColumn, and saves the file. | Show how to lock shapes in a worksheet using Aspose.Cells for .NET and then update the position of a locked shape through code. | Write an Aspose.Cells example that demonstrates shape locking and programmatic repositioning of a locked rectangle shape.
// Common Searches: Aspose.Cells C# lock shape but still move it programmatically | how to set IsLocked property for shapes in Aspose.Cells .NET | move locked rectangle shape in Excel using Aspose.Cells API | example of shape locking and repositioning with Aspose.Cells for .NET | Aspose.Cells shape lock enforcement test code
// Tags: Aspose.Cells shape locking C# | programmatic shape repositioning Aspose.Cells | Excel rectangle shape IsLocked property | Aspose.Cells create rectangle shapes | save workbook with locked shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds two rectangle shapes to the first worksheet, locks both shapes by setting IsLocked to true, programmatically moves the first shape by adjusting its UpperLeftRow and UpperLeftColumn, and saves the workbook as an .xlsx file.
class ShapeGroupLockDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add two rectangle shapes to the worksheet
            Shape shape1 = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                100, // height (pixels)
                150  // width (pixels)
            );

            Shape shape2 = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // using rectangle as fallback for older versions
                5,
                5,
                0,
                0,
                120,
                120
            );

            // Lock both shapes to prevent UI editing
            shape1.IsLocked = true;
            shape2.IsLocked = true;

            // Move the first shape programmatically
            int originalRow = shape1.UpperLeftRow;
            int originalColumn = shape1.UpperLeftColumn;
            shape1.UpperLeftRow = originalRow + 2;
            shape1.UpperLeftColumn = originalColumn + 2;

            Console.WriteLine("Shape moved programmatically:");
            Console.WriteLine($"Original Row/Column: {originalRow}/{originalColumn}");
            Console.WriteLine($"New Row/Column: {shape1.UpperLeftRow}/{shape1.UpperLeftColumn}");

            // Save the workbook to a file
            string outputPath = "ShapeGroupLockDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
