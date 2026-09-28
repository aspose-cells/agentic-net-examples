// Title: Copy a textbox shape to another worksheet while preserving size and font formatting with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that copies the first textbox shape from Sheet1 to Sheet2, preserving its position, dimensions, and all font properties using Aspose.Cells. | Show how to transfer a textbox shape between worksheets in a workbook, keeping the original size and text styling with Aspose.Cells for .NET. | Provide a C# example that clones a textbox shape from one worksheet to another, copying its text content and font attributes while maintaining layout.
// Common Searches: Aspose.Cells C# copy textbox shape from one sheet to another preserving formatting | How to duplicate a textbox shape and keep its size in Aspose.Cells .NET | Copy shape text and font attributes between worksheets using Aspose.Cells | C# Aspose.Cells example for moving a textbox while retaining layout
// Tags: copy textbox shape Aspose.Cells C# | preserve shape size and font Aspose.Cells | clone shape between worksheets .NET | textbox shape formatting Aspose.Cells | transfer shape layout Excel workbook C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel file, locates the first textbox shape on Sheet1, adds a matching textbox to Sheet2 with identical position, size, and text, copies the font properties, and saves the workbook as a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get source and destination worksheets
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            Worksheet targetSheet = workbook.Worksheets["Sheet2"];

            // Find the first textbox shape in the source sheet
            Shape sourceShape = null;
            foreach (Shape shape in sourceSheet.Shapes)
            {
                // Assume the first shape is the textbox we need
                sourceShape = shape;
                break;
            }

            if (sourceShape != null)
            {
                // Add a textbox to the target sheet with the same position and size
                Shape targetShape = targetSheet.Shapes.AddTextBox(
                    sourceShape.UpperLeftRow,
                    sourceShape.UpperLeftColumn,
                    sourceShape.Top,
                    sourceShape.Left,
                    sourceShape.Width,
                    sourceShape.Height);

                // Copy text content
                targetShape.Text = sourceShape.Text;

                // Copy font attributes
                targetShape.Font.Name = sourceShape.Font.Name;
                targetShape.Font.Size = sourceShape.Font.Size;
                targetShape.Font.IsBold = sourceShape.Font.IsBold;
                targetShape.Font.IsItalic = sourceShape.Font.IsItalic;
                targetShape.Font.Color = sourceShape.Font.Color;

                // Note: Fill and line attribute copying omitted due to API differences across versions.
            }
            else
            {
                Console.WriteLine("No shape found in the source sheet.");
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
