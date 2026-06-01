using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a formula that uses the CHAR function into cell A1 (© character)
            sheet.Cells["A1"].Formula = "=CHAR(169)";
            // Calculate the formula so the cell contains the actual character
            workbook.CalculateFormula();

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, lower right row, lower right column, width, height
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 0, 1, 0, 120, 30);

            // Link the shape to cell A1
            shape.SetLinkedCell("A1", false, false);
            // Update the shape so it reflects the linked cell's value
            shape.UpdateSelectedValue();

            // Retrieve the text displayed by the shape
            string displayedText = shape.Text;

            // Verify that the shape displays the expected special character
            Console.WriteLine($"Shape displays: {displayedText}");
            Console.WriteLine($"Verification: {(displayedText == "©" ? "Success" : "Failed")}");

            // Save the workbook
            string outputPath = "LinkedShape.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}