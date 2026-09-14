// Title: Create a rectangle shape in an Excel worksheet and lock it with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a rectangle shape to the first worksheet of a new workbook and sets its IsLocked property to true using Aspose.Cells. | Write a .NET example that creates an Excel file, inserts a locked rectangle shape with custom text, and saves it as ShapeLocked.xlsx. | Provide a C# snippet that demonstrates how to enable the locked flag on a shape added via sheet.Shapes.AddShape in Aspose.Cells.
// Common Searches: Aspose.Cells C# add rectangle shape and prevent editing | how to set IsLocked on a shape in Aspose.Cells .NET | example code to lock drawing objects in Excel using Aspose.Cells | create locked shape in Excel workbook with Aspose.Cells for .NET | C# Aspose.Cells shape lock property tutorial
// Tags: Aspose.Cells add rectangle shape C# | Aspose.Cells set shape IsLocked .NET | lock drawing object Excel Aspose.Cells | create locked shape worksheet C# | Aspose.Cells shape protection example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a new workbook, adds a rectangle shape to the first worksheet, locks the shape by setting IsLocked = true, assigns custom text, and saves the file as ShapeLocked.xlsx.
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

            // Add a rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row (zero‑based)
                2,    // upper left column (zero‑based)
                0,    // top offset in points
                0,    // left offset in points
                100,  // height in points
                200   // width in points
            );

            // Enable the locked property to prevent accidental modifications
            shape.IsLocked = true;

            // Optionally set a text for the shape
            shape.Text = "Locked Shape";

            // Save the workbook to a file
            workbook.Save("ShapeLocked.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
