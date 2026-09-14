// Title: How to ungroup an IconSet group shape and change each icon’s fill color using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells C# to locate a GroupShape named 'IconSetGroup', call Ungroup(), then set the ForeColor of each resulting Picture shape to LightGreen. | Write C# code that loads a workbook, finds the IconSet group, ungroups it, iterates over the new picture shapes, and applies a LightGreen fill using the FillFormat API.
// Common Searches: aspnet ungroup grouped shape IconSet Aspose.Cells C# example | change fill color of individual icons after ungrouping in Excel with Aspose.Cells | C# Aspose.Cells how to edit picture shapes inside an IconSet group
// Tags: Aspose.Cells ungroup GroupShape API | modify picture FillFormat in Excel using C# | IconSet group shape manipulation Aspose.Cells | C# change shape foreground color Aspose.Cells | Excel workbook shape editing with Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, finds a GroupShape named 'IconSetGroup', ungroups it so the icons become top‑level picture shapes, changes each picture's foreground fill to LightGreen via the FillFormat API, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the grouped shape named "IconSetGroup"
            GroupShape? groupShape = null;
            foreach (Shape shape in worksheet.Shapes)
            {
                if (shape.IsGroup && shape.Name == "IconSetGroup")
                {
                    groupShape = (GroupShape)shape;
                    break;
                }
            }

            if (groupShape != null)
            {
                try
                {
                    // Ungroup the shape; after this call the child shapes become top‑level shapes
                    groupShape.Ungroup();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to ungroup shape: {ex.Message}");
                }

                // Modify each picture shape (icon) that resulted from ungrouping
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Check if the shape is a picture
                    if ((MsoDrawingType)shape.Type == MsoDrawingType.Picture)
                    {
                        // Change the foreground fill color to LightGreen
                        shape.FillFormat.ForeColor = Color.LightGreen;
                    }
                }
            }
            else
            {
                Console.WriteLine("Group shape 'IconSetGroup' not found.");
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
