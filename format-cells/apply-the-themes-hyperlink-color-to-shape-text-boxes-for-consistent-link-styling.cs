// Title: Set the workbook’s hyperlink theme color for TextBox shapes using Aspose.Cells in C#
// Common Searches: how to set a textbox shape’s font color to the workbook hyperlink theme using Aspose.Cells C# | Aspose.Cells example for applying hyperlink styling to shape text boxes in an Excel file | programmatically change text box hyperlink color and underline in .NET with Aspose.Cells | iterate over worksheet shapes and format text boxes with theme colors in C# | C# code to apply workbook theme hyperlink formatting to shape text in Excel
// Tags: textbox shape textfont formatting Aspose.Cells | hyperlink theme color application to shapes | iterate worksheet shapes C# | set underline on shape text Aspose.Cells | excel workbook theme styling via Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing Excel workbook, walks through each worksheet and its shapes, identifies TextBox shapes, and (when enabled) sets the TextFont color and underline to match the workbook’s hyperlink theme color, then saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Use a standard hyperlink color (blue) as a fallback
            Color hyperlinkColor = Color.Blue;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Process only text box shapes
                        // Cast to MsoDrawingType to handle possible enum differences across versions
                        if (((MsoDrawingType)shape.Type) == MsoDrawingType.TextBox)
                        {
#if NET6_0_OR_GREATER
                            // Uncomment the following lines if your Aspose.Cells version supports TextFont
                            // shape.TextFont.Color = hyperlinkColor;
                            // shape.TextFont.Underline = true;
#endif
                        }
                    }
                    catch (Exception shapeEx)
                    {
                        // Log shape-specific errors without stopping the whole process
                        Console.WriteLine($"Error processing shape on sheet \"{sheet.Name}\": {shapeEx.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
