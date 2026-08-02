// Title: C# – Move Excel Shapes to Front or Back and Retrieve Z‑Order with Aspose.Cells
// Description: This example creates a workbook, adds two overlapping rectangle shapes, displays their initial ZOrderPosition, moves the first shape to the front with ToFrontOrBack(1), moves the second shape to the back with ToFrontOrBack(-1), shows the updated ZOrderPosition values, and saves the file as ShapeZOrderDemo.xlsx.
// Keywords: Aspose.Cells shape ZOrderPosition | ToFrontOrBack method | bring shape to front Aspose.Cells | send shape to back Aspose.Cells | Excel shape layering .NET | C# Aspose.Cells shape ordering
// Common Searches: Aspose.Cells change shape Z‑order C# | How to bring a shape to front in Excel using Aspose.Cells | Retrieve ZOrderPosition after moving a shape Aspose.Cells | Send shape to back Aspose.Cells .NET | Shape layering example Aspose.Cells
// Developer Intent: The developer needs to programmatically adjust the Z‑order of worksheet shapes and read the resulting ZOrderPosition values for verification or further processing.
// Use Cases: Ensure a comment or annotation shape always appears above other objects before exporting the worksheet. | Reorder overlapping images in a financial report to match a required visual hierarchy. | Log Z‑order changes during automated worksheet generation to debug layout problems.
// AI Prompts: Generate C# code that sets specific Z‑order values for a collection of shapes using Aspose.Cells. | Show how to reset all shapes to their original ZOrderPosition after modifications in a workbook. | Provide a loop that iterates through every shape in a worksheet and prints its ZOrderPosition.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.IO;

// This example creates a workbook, adds two overlapping rectangle shapes, displays their initial ZOrderPosition, moves the first shape to the front with ToFrontOrBack(1), moves the second shape to the back with ToFrontOrBack(-1), shows the updated ZOrderPosition values, and saves the file as ShapeZOrderDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two overlapping rectangle shapes
            Shape shape1 = worksheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
            Shape shape2 = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

            // Record initial Z-order positions
            Console.WriteLine("Initial ZOrderPosition - shape1: " + shape1.ZOrderPosition);
            Console.WriteLine("Initial ZOrderPosition - shape2: " + shape2.ZOrderPosition);

            // Bring shape1 to the front (positive order value)
            shape1.ToFrontOrBack(1);
            Console.WriteLine("After bringing shape1 to front - shape1 ZOrderPosition: " + shape1.ZOrderPosition);
            Console.WriteLine("After bringing shape1 to front - shape2 ZOrderPosition: " + shape2.ZOrderPosition);

            // Send shape2 to the back (negative order value)
            shape2.ToFrontOrBack(-1);
            Console.WriteLine("After sending shape2 to back - shape2 ZOrderPosition: " + shape2.ZOrderPosition);
            Console.WriteLine("After sending shape2 to back - shape1 ZOrderPosition: " + shape1.ZOrderPosition);

            // Save the workbook
            string outputPath = "ShapeZOrderDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to " + Path.GetFullPath(outputPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
