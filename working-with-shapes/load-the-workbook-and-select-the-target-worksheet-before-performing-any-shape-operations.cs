// Title: Load an Excel workbook, select a specific worksheet, and insert a rectangle shape using Aspose.Cells for .NET
// AI Prompts: Load a workbook from a file path, choose a worksheet by its name, add a rectangle shape at row 1 column 0 with width 100 px and height 50 px, then save the workbook. | Change the code to select the first worksheet by index and insert an ellipse shape instead of a rectangle while keeping the same dimensions. | Write a reusable method that takes a file path, worksheet identifier (name or index), shape type, and size parameters, adds the shape to the specified sheet, and returns the updated workbook.
// Common Searches: Aspose.Cells C# load workbook and add shape to a specific sheet | select worksheet by name before inserting drawing with Aspose.Cells .NET | add rectangle shape to Excel sheet using Aspose.Cells drawing API example | how to change shape type to ellipse when adding to a worksheet with Aspose.Cells | save modified workbook after adding shapes using Aspose.Cells C#
// Tags: load workbook select worksheet Aspose.Cells | add rectangle shape to Excel sheet C# | worksheet shape insertion using Aspose.Cells drawing API | select worksheet by name before shape operation | Aspose.Cells shape type parameter example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads 'input.xlsx' into a Workbook, selects the worksheet named 'Sheet1', adds a rectangle shape with specified pixel dimensions to that sheet, and saves the result as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Select the target worksheet (by name or index)
        Worksheet worksheet = workbook.Worksheets["Sheet1"]; // replace with your sheet name
        // Worksheet worksheet = workbook.Worksheets[0]; // alternative: select by index

        // Example shape operation: add a rectangle shape to the selected worksheet
        Shape rectangle = worksheet.Shapes.AddShape(
            MsoDrawingType.Rectangle, // shape type
            1,   // upper left row
            0,   // upper left column
            1,   // top offset in pixels
            0,   // left offset in pixels
            100, // width in pixels
            50   // height in pixels
        );

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
