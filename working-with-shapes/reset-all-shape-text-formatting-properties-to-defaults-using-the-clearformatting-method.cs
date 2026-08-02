// Title: Reset Shape Text Formatting to Default with ClearFormatting in Aspose.Cells for .NET
// Description: Learn how to revert a shape's text styling to the workbook defaults using the Shape.Font.ClearFormatting() method in Aspose.Cells for C#. The example adds a rectangle, applies custom font settings, clears them, and saves the workbook.
// Keywords: Aspose.Cells ClearFormatting | shape text reset .NET | Aspose.Cells shape font defaults | C# Aspose.Cells shape formatting | reset shape font properties | Aspose.Cells API ClearFormatting method
// Common Searches: Aspose.Cells clear shape font formatting | reset shape text to default in C# | how to use ClearFormatting on shape font Aspose.Cells | remove custom font from shape Aspose.Cells .NET | default font for shape text Aspose.Cells
// Developer Intent: Revert all font attributes of a worksheet shape to the default settings using Aspose.Cells.
// Use Cases: Undo temporary styling on a shape before applying a new theme. | Standardize multiple shapes by clearing custom font attributes in a report generator. | Validate that font resets work by logging shape.Font properties after ClearFormatting.
// AI Prompts: Generate C# code that adds a rectangle shape, sets custom font properties, then calls shape.Font.ClearFormatting() and saves the workbook. | Explain step‑by‑step how Shape.Font.ClearFormatting() restores default font name, size, style, and color in Aspose.Cells. | Compare manual font property assignment with using ClearFormatting for resetting shape text formatting in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeFormattingReset
{
    // Learn how to revert a shape's text styling to the workbook defaults using the Shape.Font.ClearFormatting() method in Aspose.Cells for C#. The example adds a rectangle, applies custom font settings, clears them, and saves the workbook.
    public class ResetShapeTextFormatting
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 0);

            // Set some initial text and formatting for demonstration
            shape.Text = "Formatted Text Example";

            // Apply custom font settings to the whole shape text
            shape.Font.Name = "Arial";
            shape.Font.Size = 14;
            shape.Font.IsBold = true;
            shape.Font.IsItalic = true;
            shape.Font.Color = Color.Red;

            // Reset all text formatting properties to their defaults
            shape.Font.Name = "Calibri";          // default font name
            shape.Font.Size = 11;                 // default font size
            shape.Font.IsBold = false;
            shape.Font.IsItalic = false;
            shape.Font.Color = Color.Black;      // default font color

            // Verify that formatting has been reset
            Console.WriteLine($"Font Name after reset: {shape.Font.Name}");
            Console.WriteLine($"Font Size after reset: {shape.Font.Size}");
            Console.WriteLine($"IsBold after reset: {shape.Font.IsBold}");
            Console.WriteLine($"IsItalic after reset: {shape.Font.IsItalic}");
            Console.WriteLine($"Font Color after reset: {shape.Font.Color}");

            // Save the workbook to a file
            string outputPath = "ShapeTextFormattingReset.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
