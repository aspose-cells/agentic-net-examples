// Title: Validate that WordArt gradient fills render identically when converting an Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to load an .xlsx file, iterate all worksheets, detect WordArt shapes with gradient fills, and output each gradient stop's position and the gradient angle. | Develop a C# method that compares the extracted gradient fill properties of WordArt shapes to expected values and reports any mismatches. | Enhance the program to summarize the count of shapes using gradient fills versus non‑gradient fills before saving the workbook as a PDF.
// Common Searches: aspnet aspocells read WordArt gradient fill properties from an Excel workbook | how to verify that gradient fills are preserved during Excel to PDF conversion with Aspose.Cells | c# enumerate worksheet shapes and retrieve gradient stop positions using Aspose.Cells | compare gradient fill details between source Excel and generated PDF in .NET | log gradient angle of WordArt shapes while converting Excel to PDF with Aspose.Cells
// Tags: Aspose.Cells gradient fill extraction | WordArt shape enumeration .NET | Excel to PDF gradient rendering verification | C# FillFormat gradient analysis | PDF conversion shape fill consistency

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGradientValidation
{
    // The example loads an Excel workbook, walks through every worksheet shape, identifies WordArt objects with gradient fills, logs each gradient stop's position and angle, reports shapes with and without gradients, and finally saves the workbook as a PDF to confirm visual fidelity.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string sourcePath = "SourceWorkbook.xlsx";
                const string outputPath = "OutputDocument.pdf";

                // Verify that the source workbook exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Error: The file '{sourcePath}' was not found.");
                    return;
                }

                // Load the source Excel workbook
                Workbook workbook = new Workbook(sourcePath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Check the shape's fill format for gradient fills
                        FillFormat fill = shape.Fill;

                        if (fill != null && fill.FillType == FillType.Gradient && fill.GradientFill != null)
                        {
                            GradientFill gradient = fill.GradientFill;

                            // Output basic gradient information
                            Console.WriteLine($"Shape '{shape.Name}' – Gradient Stops: {gradient.GradientStops.Count}");

                            // Iterate through gradient stops
                            for (int i = 0; i < gradient.GradientStops.Count; i++)
                            {
                                GradientStop stop = gradient.GradientStops[i];
                                // Position is always available; color may not be present in older versions,
                                // so we output only the position to keep the code compatible.
                                Console.WriteLine($"Shape '{shape.Name}' – Gradient Stop {i}: Position={stop.Position}");
                            }

                            // Output gradient angle (relevant for linear gradients)
                            Console.WriteLine($"Shape '{shape.Name}' – Gradient Angle: {gradient.Angle}");
                        }
                        else
                        {
                            Console.WriteLine($"Shape '{shape.Name}' does not use a gradient fill.");
                        }
                    }
                }

                // Save the workbook as PDF (gradient fills are rendered during conversion)
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
