// Title: How to bring a rectangle shape to the front layer in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new Workbook, add a rectangle shape to the first worksheet, call shape.ToFrontOrBack(1) to set its Z‑order to the front, and save the file as .xlsx. | Use the Aspose.Cells C# API to insert a rectangle shape and move it to the top layer by passing a positive integer to the ToFrontOrBack method before saving the workbook.
// Common Searches: Aspose.Cells C# move shape to front layer using ToFrontOrBack | How to set Z-order of a shape in Excel with Aspose.Cells .NET | Example of bringing a rectangle shape to the top in an Aspose.Cells workbook | Using ToFrontOrBack method with a positive integer in Aspose.Cells C# | Shape ordering in Excel files with Aspose.Cells API
// Tags: Aspose.Cells shape ToFrontOrBack | C# Aspose.Cells rectangle shape Z-order | Excel .xlsx shape front layer | Aspose.Cells add shape front order | C# move shape to top Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a rectangle shape to the first worksheet, moves the shape to the front layer with ToFrontOrBack(1), and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
        var shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

        // Move the shape to the front layer using a positive integer argument
        shape.ToFrontOrBack(1); // any positive integer brings the shape to the front

        // Save the workbook (save rule)
        workbook.Save("output.xlsx");
    }
}
