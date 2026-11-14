// Title: Create a rectangle shape linked to a cell with a TEXT formula and custom number format using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape text to =TEXT(A1,"Qty:" 0.00) formula | How to display a custom formatted number inside an Excel shape with Aspose.Cells | Link rectangle shape to a worksheet cell using TEXT function in .NET | Calculate shape text after assigning a formula in an Aspose.Cells workbook | Save workbook after adding a shape that references a cell in C# Aspose.Cells
// Tags: rectangle shape with TEXT formula Aspose.Cells | custom formatted number in shape text C# | shape linked to worksheet cell Aspose.Cells | evaluate shape text after formula calculation | add and save shape with cell reference .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes a numeric value to cell A1, adds a rectangle shape, assigns its Text property to a TEXT formula that prefixes the value with "Qty:" and formats it to two decimals, forces formula recalculation so the shape displays the evaluated text, prints the result, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into cell A1
            Cell targetCell = sheet.Cells["A1"];
            targetCell.PutValue(123.456);

            // Add a rectangle shape to the worksheet
            // Parameters: MsoDrawingType, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 50, 50, 100, 200);

            // Link the shape's text to the cell using TEXT function with custom format
            // The format string includes custom text "Qty:" followed by the number with two decimals
            shape.Text = "=TEXT(A1,\"\\\"Qty:\\\" 0.00\")";

            // Recalculate formulas so the shape displays the evaluated text
            workbook.CalculateFormula();

            // Verify the shape's displayed text (after calculation this holds the evaluated result)
            string displayedText = shape.Text;
            Console.WriteLine("Shape text after evaluation: " + displayedText);

            // Save the workbook
            string outputPath = "output.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
