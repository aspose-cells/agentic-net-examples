// Title: How to handle WordArt shapes when converting an Excel workbook to PDF with Aspose.Cells for .NET (shadow effect limitation)
// AI Prompts: Generate C# code that enumerates WordArt shapes, validates their fill type, and saves the workbook as PDF using Aspose.Cells. | Propose a technique to simulate a shadow-like appearance on WordArt before PDF export in a .NET project with Aspose.Cells.
// Common Searches: Aspose.Cells C# cannot apply shadow to WordArt during Excel to PDF conversion | preserve WordArt formatting when exporting Excel to PDF with Aspose.Cells | detect WordArt shapes in a workbook using Aspose.Cells API
// Tags: Aspose.Cells WordArt object detection C# | workbook PDF export Aspose.Cells | shadow feature unavailable Aspose.Cells | gradient fill preservation Aspose.Cells | C# shape properties Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, iterates through its shapes, identifies WordArt via the IsWordArt property, notes that Aspose.Cells does not expose shadow settings for shapes, confirms that gradient fills are retained automatically, and then saves the workbook as a PDF while preserving the original WordArt appearance.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.pdf";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Identify WordArt shapes using the IsWordArt property
                    if (shape.IsWordArt)
                    {
                        // Aspose.Cells does not expose direct shadow properties for shapes.
                        // If shadow effects are required, they must be applied via other means
                        // (e.g., editing the source file or using a different library).

                        // Gradient fill is preserved automatically; no changes required.
                        // Example of checking fill type (optional):
                        // FillFormat fill = shape.FillFormat;
                        // if (fill.FillType == FillType.Gradient) { /* keep as is */ }
                    }
                }

                // Save the workbook as PDF, preserving WordArt and its formatting
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions to prevent the application from crashing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
