// Title: Export an Excel workbook with WordArt objects to PDF while enumerating and logging each WordArt gradient fill using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, walks through every worksheet, finds WordArt shapes, prints each gradient stop position, and saves the whole workbook as a single PDF. | Adapt the example to generate a separate PDF for each worksheet that contains at least one WordArt shape with a gradient fill, ensuring the gradients remain visible. | Add robust exception handling that reports missing input files, shape‑processing errors, and PDF‑saving failures with detailed diagnostic messages.
// Common Searches: how to export Excel WordArt gradients to PDF using Aspose.Cells C# | enumerate WordArt shapes and read gradient stops in a .xlsx with Aspose.Cells | verify that WordArt gradient fills are preserved after converting to PDF in .NET | C# Aspose.Cells save each worksheet containing WordArt as separate PDF files | debug gradient fill information for WordArt objects in an Excel workbook
// Tags: Aspose.Cells WordArt gradient inspection | C# shape collection gradient stops Aspose | PDF generation preserving Excel shape fills | worksheet-specific PDF output Aspose.Cells | gradient fill verification Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel file with Aspose.Cells, iterates through all worksheets and their shape collections, identifies WordArt objects, logs gradient fill details such as stop count and positions, and finally saves the workbook as a PDF while handling file‑access and shape‑processing errors.
class Program
{
    static void Main()
    {
        string inputPath = "InputWithWordArt.xlsx";
        string outputPath = "OutputWithWordArt.pdf";

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

            // Iterate through worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Process only WordArt shapes
                    if (shape.IsWordArt)
                    {
                        try
                        {
                            // Get the fill format of the shape
                            FillFormat fill = shape.Fill;

                            // Check if the fill is a gradient
                            if (fill.FillType == FillType.Gradient && fill.GradientFill != null)
                            {
                                Console.WriteLine($"Worksheet: {sheet.Name}, WordArt Index: {i}");
                                Console.WriteLine($"  Gradient Stops Count: {fill.GradientFill.GradientStops.Count}");

                                // List each gradient stop (color property omitted due to API version differences)
                                for (int j = 0; j < fill.GradientFill.GradientStops.Count; j++)
                                {
                                    GradientStop stop = fill.GradientFill.GradientStops[j];
                                    Console.WriteLine($"    Stop {j}: Position={stop.Position}");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Worksheet: {sheet.Name}, WordArt Index: {i} does NOT have a gradient fill.");
                            }
                        }
                        catch (Exception exShape)
                        {
                            Console.WriteLine($"Error processing shape at index {i} on sheet {sheet.Name}: {exShape.Message}");
                        }
                    }
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            try
            {
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Error saving PDF: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
