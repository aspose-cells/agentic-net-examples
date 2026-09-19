// Title: Set the Z‑order of a TextBox shape to bring it to the front in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Add a line that sets the newly created TextBox's ZOrder property so it appears above all other shapes in the worksheet. | Modify the sample to call the Shape.ZOrder method (or assign ZOrder = 0) after adding the TextBox, ensuring the textbox is rendered on the top layer.
// Common Searches: Aspose.Cells C# change shape Z-order | bring textbox to front of other shapes in Excel using Aspose.Cells | set shape layering order in a .NET workbook with Aspose.Cells
// Tags: Aspose.Cells shape ZOrder property | C# bring textbox to front Excel | set shape layering Aspose.Cells .NET | adjust shape order worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, adds a TextBox shape, and saves the file, but it does not modify the shape's Z‑order. To ensure the textbox appears above any overlapping objects, set its ZOrder property (or use the appropriate method) before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, row offset, column offset, width (pixels), height (pixels)
            TextBox textBox = sheet.Shapes.AddTextBox(2, 1, 0, 0, 150, 60);
            textBox.Text = "Hello Aspose.Cells";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
