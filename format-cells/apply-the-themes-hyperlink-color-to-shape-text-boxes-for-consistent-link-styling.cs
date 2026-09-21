// Title: Set TextBox AutoShape font color to the workbook's hyperlink theme color using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to iterate all worksheets and change the font color of every TextBox AutoShape to the workbook's hyperlink theme color. | Apply the default hyperlink blue color to shape text boxes in an Excel file with Aspose.Cells while preserving the workbook's theme.
// Common Searches: c# aspose.cells change textbox shape font to hyperlink color | how to apply Excel hyperlink theme color to AutoShape text using Aspose.Cells | set shape text color based on workbook theme in Aspose.Cells .NET | update TextBox AutoShape font color to match hyperlink style programmatically
// Tags: Aspose.Cells set textbox font color | hyperlink theme color for shape text | iterate worksheets shapes C# | apply workbook theme to AutoShape text | Excel hyperlink style Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, iterates through each worksheet and its shapes, and sets the font color of any TextBox AutoShape to the workbook's hyperlink theme color (default blue) before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Default hyperlink color (blue)
            Color hyperlinkColor = Color.Blue;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only TextBox AutoShape types
                    if (shape.AutoShapeType == AutoShapeType.TextBox)
                    {
                        // Apply the hyperlink color to the shape's text font
                        var font = shape.Font;
                        if (font != null)
                        {
                            font.Color = hyperlinkColor;
                        }
                    }
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
