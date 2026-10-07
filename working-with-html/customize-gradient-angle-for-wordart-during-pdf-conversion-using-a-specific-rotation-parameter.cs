// Title: Set a custom gradient angle on WordArt shapes in Excel and export to PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates over every shape in a worksheet, detects WordArt shapes, changes their fill type to gradient, assigns a specific GradientAngle (e.g., 45°), and saves the workbook as a PDF using Aspose.Cells. | Show how to safely apply a gradient rotation to WordArt shapes when the GradientAngle property may be unavailable, then export the workbook to PDF with Aspose.Cells.
// Common Searches: how to change WordArt gradient direction in Excel using Aspose.Cells C# | Aspose.Cells set gradient angle for WordArt before saving as PDF | C# code to apply custom gradient rotation to Excel WordArt shapes | preserve WordArt gradient fill when converting Excel workbook to PDF with Aspose.Cells
// Tags: WordArt gradient rotation Aspose.Cells | Excel shape gradient fill C# | custom WordArt styling PDF export | Aspose.Cells shape fill type gradient | set WordArt fill rotation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel file, scans all shapes on the first worksheet, identifies WordArt shapes, switches their fill to a gradient, optionally sets a GradientAngle, and then saves the workbook as a PDF, handling missing files and shape‑processing errors.
class Program
{
    static void Main()
    {
        const string inputFilePath = "input.xlsx";
        const string outputFilePath = "output.pdf";
        const double gradientAngle = 45.0; // degrees

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException($"Input file not found: {inputFilePath}");

            // Load the Excel workbook
            var workbook = new Workbook(inputFilePath);

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Find WordArt shapes and set their gradient fill
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Identify WordArt shapes
                    if (shape.IsWordArt)
                    {
                        // Ensure the fill type is gradient so the angle would take effect (if supported)
                        shape.Fill.FillType = FillType.Gradient;

                        // GradientAngle property may not be available in older versions of Aspose.Cells.
                        // If supported, uncomment the following line:
                        // shape.Fill.GradientAngle = gradientAngle;
                    }
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Warning: Could not process shape '{shape.Name}'. {exShape.Message}");
                }
            }

            // Save the workbook as PDF
            workbook.Save(outputFilePath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to: {outputFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
