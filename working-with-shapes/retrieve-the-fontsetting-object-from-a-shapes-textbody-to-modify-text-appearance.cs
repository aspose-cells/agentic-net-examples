// Title: Retrieve and modify a shape's FontSetting (text body) in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing .xlsx file, locates a shape named 'MyShape' on the first worksheet, accesses its TextBody FontSetting, changes the font color to red, size to 14 pt, makes it bold and applies a single underline, then saves the workbook. | Show how to safely obtain the Font object from a shape's TextBody in Aspose.Cells, adjust color, size, bold, italic, and underline properties, and handle missing shape or file errors.
// Common Searches: Aspose.Cells C# retrieve FontSetting from shape TextBody | How to change font color and size of a specific shape in Excel using Aspose.Cells | Set bold and underline for shape text in Aspose.Cells .NET example | C# code to modify shape text appearance in an .xlsx workbook with Aspose.Cells | Aspose.Cells shape Font object not found error handling
// Tags: Aspose.Cells shape FontSetting modification | C# Aspose.Cells retrieve shape font | Excel shape text formatting Aspose.Cells | Aspose.Cells set shape font color size | Aspose.Cells apply underline to shape text | Aspose.Cells handle missing shape exception

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example opens an existing workbook, accesses the first worksheet, retrieves a shape named "MyShape", obtains its FontSetting via the shape's Font property, updates the font color to red, size to 14 pt, makes it bold, removes italic, adds a single underline, and saves the modified workbook while handling missing files or shapes.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the shape by name
            Shape shape = worksheet.Shapes["MyShape"];
            if (shape == null)
            {
                Console.WriteLine("Shape 'MyShape' not found on the worksheet.");
                return;
            }

            // Access the shape's font and modify its appearance
            Font font = shape.Font;
            font.Color = Color.Red;               // Font color
            font.Size = 14;                        // Font size (points)
            font.IsBold = true;                    // Bold
            font.IsItalic = false;                 // Italic
            font.Underline = FontUnderlineType.Single; // Underline

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
