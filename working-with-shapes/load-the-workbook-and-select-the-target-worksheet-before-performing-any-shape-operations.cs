// Title: Add a Rectangle Shape to a Worksheet using Aspose.Cells for .NET
// Description: Demonstrates how to load or create an Excel workbook, select a worksheet, insert a rectangle shape, retrieve the shape's parent worksheet via the Shape.Worksheet property, display the worksheet name, and save the file.
// Keywords: Aspose.Cells add rectangle shape | Shape.Worksheet property | load workbook Aspose.Cells .NET | create shape on worksheet | retrieve parent worksheet from shape
// Common Searches: Aspose.Cells add rectangle to worksheet | how to get worksheet of a shape Aspose.Cells | load Excel file and insert shapes C# | Shape.Worksheet example Aspose.Cells | create shape in new workbook Aspose.Cells
// Developer Intent: Load or create a workbook, select a specific worksheet, add a rectangle shape, and determine which worksheet owns the shape.
// Use Cases: Programmatically add visual markers (e.g., rectangles) to a worksheet for reporting. | Log the parent worksheet name of dynamically created shapes for audit trails. | Build utilities that need to verify or modify shapes based on their containing worksheet.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, adds a circle shape to the second worksheet, and returns the worksheet name of the shape. | Show how to create multiple shapes across different worksheets and collect each shape's parent worksheet name into a list using Aspose.Cells for .NET. | Write a function that checks for a workbook file, creates it if missing, adds a triangle shape to a given worksheet index, and saves the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeDemo
{
    // Demonstrates how to load or create an Excel workbook, select a worksheet, insert a rectangle shape, retrieve the shape's parent worksheet via the Shape.Worksheet property, display the worksheet name, and save the file.
    public class ShapeWorksheetExample
    {
        // Entry point required by the runtime
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Ensure the input file exists; otherwise create a new workbook
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Select the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

            // Retrieve the worksheet that contains the shape
            Worksheet shapeWorksheet = shape.Worksheet;

            // Output the name of the worksheet that owns the shape
            Console.WriteLine("Shape belongs to worksheet: " + shapeWorksheet.Name);

            // Save the workbook with the new shape
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
    }
}
