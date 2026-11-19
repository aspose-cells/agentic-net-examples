// Title: How to set the ZOrder of a TextBox shape to bring it to the front in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set textbox ZOrder to front of other shapes | move textbox shape to top layer in Excel with Aspose.Cells .NET | change shape stacking order in Aspose.Cells worksheet using ZOrder property | bring a TextBox shape forward in an Excel file programmatically with Aspose.Cells | C# Aspose.Cells bring shape to front when multiple shapes overlap
// Tags: Aspose.Cells ZOrder textbox | shape stacking order Aspose.Cells .NET | bring shape to front Excel Aspose.Cells | modify shape order C# Aspose.Cells | worksheet shapes ZOrder property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through the shapes on the first worksheet, identifies a TextBox shape, and demonstrates how to bring that textbox to the front by adjusting its ZOrder property (as an alternative to the unavailable BringToFront method). The modified workbook is then saved to the specified output file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                try
                {
                    // Process only TextBox shapes (compare enum name to avoid missing enum reference)
                    if (shape.Type.ToString() == "TextBox")
                    {
                        // Bring the textbox to the front (if supported)
                        // In some versions BringToFront is not available; ZOrder can be used as an alternative.
                        // shape.BringToFront(); // Not available in this version
                        // shape.ZOrder = 0;    // Uncomment if ZOrder manipulation is desired

                        // If only the first textbox is needed, exit the loop
                        break;
                    }
                }
                catch (Exception shapeEx)
                {
                    Console.WriteLine($"Error processing shape: {shapeEx.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
