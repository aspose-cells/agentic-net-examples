// Title: Create a rectangle shape linked to a cell with the TEXT function for percentage display and verify it using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link shape text to a cell using TEXT function | display cell percentage inside a shape with Aspose.Cells .NET | verify shape text shows formatted percent value in Aspose.Cells | add rectangle shape and bind it to A1 formatted as 0.00% using Aspose.Cells | calculate formulas for shape text after linking to a cell in Aspose.Cells
// Tags: link shape text to cell TEXT function Aspose.Cells | percentage format in shape text Aspose.Cells | add rectangle shape to worksheet Aspose.Cells | apply built‑in 0.00% number format Aspose.Cells | programmatic verification of shape text value Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes 0.256 to cell A1, applies a 0.00% number format, adds a rectangle shape, sets its text to =TEXT(A1,"0.00%"), forces formula calculation, checks that the shape text equals "25.60%", and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a numeric value in cell A1 (0.256 = 25.6%)
            Cell targetCell = cells["A1"];
            targetCell.PutValue(0.256);

            // Apply percentage format to the cell (e.g., 0.00%)
            Style percentStyle = workbook.CreateStyle();
            percentStyle.Number = 10; // Built‑in format for "0.00%"
            targetCell.SetStyle(percentStyle);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type enum
                2, 0, 5, 5, 100, 30);

            // Link the shape's text to the cell using the TEXT function for percentage display
            shape.Text = "=TEXT(A1,\"0.00%\")";

            // Force formula calculation so the shape's text is updated
            workbook.CalculateFormula();

            // Retrieve the displayed text from the shape
            string displayedText = shape.Text;

            // Verify that the shape shows the expected percentage
            Console.WriteLine("Shape text: " + displayedText);
            if (displayedText == "25.60%")
            {
                Console.WriteLine("Percent display verified successfully.");
            }
            else
            {
                Console.WriteLine("Percent display verification failed.");
            }

            // Save the workbook (lifecycle rule: save)
            string outputPath = "PercentageShapeDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
