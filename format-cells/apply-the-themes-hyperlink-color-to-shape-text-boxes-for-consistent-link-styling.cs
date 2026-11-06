// Title: Set TextBox shape font color to the workbook theme's hyperlink color using Aspose.Cells for .NET (C#)
// Common Searches: how to apply workbook theme hyperlink color to text boxes in an Excel file with Aspose.Cells C# | Aspose.Cells C# change font color of all TextBox shapes to the theme hyperlink color | retrieve Excel theme hyperlink color in .NET and use it for shape text formatting | set text box font color based on workbook theme using Aspose.Cells for .NET | apply theme hyperlink color to shape fonts in a generated spreadsheet with Aspose.Cells
// Tags: Aspose.Cells set shape font color | Excel theme hyperlink color extraction .NET | TextBox shape formatting Aspose.Cells | apply workbook theme colors to shapes | C# reflection workbook theme properties Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, obtains the Hyperlink color defined in the workbook's theme (falling back to blue if unavailable), iterates through each worksheet and its shapes, and for every TextBox shape sets the Font.Color to the retrieved theme hyperlink color before saving the updated file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Default hyperlink color (fallback)
            Color hyperlinkColor = Color.Blue;

            // Attempt to obtain the hyperlink color from the workbook's theme (if supported)
            try
            {
                var themeProp = workbook.GetType().GetProperty("Theme");
                if (themeProp != null)
                {
                    var themeObj = themeProp.GetValue(workbook);
                    var hyperlinkProp = themeObj?.GetType().GetProperty("Hyperlink");
                    if (hyperlinkProp != null)
                    {
                        var colorObj = hyperlinkProp.GetValue(themeObj);
                        if (colorObj is Color c)
                            hyperlinkColor = c;
                    }
                }
            }
            catch
            {
                // Keep default color if reflection fails
            }

            // Iterate through worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Process only TextBox shapes
                        if (shape.Type == AutoShapeType.TextBox)
                        {
                            // Apply the hyperlink color to the text inside the TextBox
                            if (shape.Font != null)
                                shape.Font.Color = hyperlinkColor;
                        }
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Error processing shape on sheet '{sheet.Name}': {exShape.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Error saving workbook: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
