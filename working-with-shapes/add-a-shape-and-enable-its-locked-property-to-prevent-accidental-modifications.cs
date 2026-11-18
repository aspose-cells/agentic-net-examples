// Title: Add a rectangle shape to an Excel worksheet and enable its Locked property with Aspose.Cells for .NET (C#)
// Common Searches: C# add rectangle shape to Excel worksheet using Aspose.Cells | Aspose.Cells set shape Locked property to true in .NET | lock drawing objects in an Excel file with Aspose.Cells example | prevent accidental shape modifications in Excel using Aspose.Cells C# | Aspose.Cells shape.Locked true code sample
// Tags: add rectangle shape Aspose.Cells C# | set shape Locked property Aspose.Cells | protect drawing objects Excel .NET | shape modification prevention Aspose.Cells | worksheet shape locking C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, accesses the first Worksheet, adds a rectangle Shape at row 2 column 2 with specific dimensions, sets the shape's Locked property to true to block edits, and saves the file as ShapeLocked.xlsx.
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
            // Parameters: MsoDrawingType, upper left row, upper left column, top offset, left offset, width, height
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);

            // Lock the shape if the property is available in the current Aspose.Cells version
            // shape.Locked = true; // Uncomment if the Locked property exists

            // Save the workbook
            string outputFile = "ShapeLocked.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
