// Title: Add a rectangle shape linked to a cell and display its value as a percentage using the TEXT function in Aspose.Cells for .NET
// AI Prompts: Insert a rectangle shape on a worksheet, assign its Text property to =TEXT(A1,"0.00%"), and recalculate formulas with Aspose.Cells. | Bind a shape to cell A1 and show the cell's numeric value formatted as a percentage inside the shape using C#. | Save the workbook after the shape displays the formatted percentage and verify the output.
// Common Searches: Aspose.Cells how to show cell value inside a shape as percentage | C# set shape text to formula using TEXT function Aspose.Cells | link rectangle shape to cell and format with custom number format Aspose.Cells | calculate shape text after adding formula in Aspose.Cells .NET
// Tags: add rectangle shape Aspose.Cells | shape text formula TEXT function | percentage custom number format shape | calculate formulas for shape content | save workbook with shape text Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a workbook, writes 0.1234 to A1, adds a rectangle shape, sets its Text to =TEXT(A1,"0.00%"), recalculates formulas so the shape shows 12.34%, prints the result, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Place a numeric value in cell A1 (e.g., 0.1234)
            sheet.Cells["A1"].PutValue(0.1234);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape;
            try
            {
                shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    1, 0, 1, 0, 100, 30);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to add shape: {ex.Message}");
                return;
            }

            // Set the shape's text to a formula that formats the cell value as a percentage
            shape.Text = "=TEXT(A1,\"0.00%\")";

            // Calculate all formulas in the workbook so the shape text is evaluated
            workbook.CalculateFormula();

            // Retrieve the displayed text from the shape after calculation
            string displayedText = shape.Text;

            // Verify the result by printing it to the console (expected output: 12.34%)
            Console.WriteLine("Shape text: " + displayedText);

            // Optional: save the workbook to a file
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
