// Title: Add a rectangle shape to an Excel worksheet and lock it using Aspose.Cells for .NET (C#)
// AI Prompts: Generate an Excel workbook, insert a rectangle shape on the first sheet, and set its IsLocked flag to true with Aspose.Cells in C#. | Create a worksheet, add a rectangle drawing, and lock the shape to prevent edits using the Aspose.Cells API.
// Common Searches: C# Aspose.Cells how to lock a shape after adding it | example of adding a rectangle shape to a worksheet and making it read‑only with Aspose.Cells | Aspose.Cells set shape IsLocked property to prevent editing | lock Excel shape programmatically using Aspose.Cells .NET | prevent accidental modification of shapes in generated Excel workbook Aspose.Cells
// Tags: add rectangle shape Aspose.Cells | lock shape IsLocked Aspose.Cells | shape protection Excel C# | worksheet shape locking Aspose.Cells | prevent shape editing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code creates a new workbook, adds a rectangle shape to the first worksheet, locks the shape by setting its IsLocked property to true, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper-left row, upper-left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // Lock the shape to prevent accidental modifications
            shape.IsLocked = true;

            // Define output file path
            string outputPath = "output.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
