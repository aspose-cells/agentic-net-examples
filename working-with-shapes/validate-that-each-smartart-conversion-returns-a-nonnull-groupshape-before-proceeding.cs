// Title: Validate SmartArt‑to‑GroupShape conversion in Aspose.Cells for .NET
// Description: Loads an Excel file, verifies its existence, scans every worksheet for SmartArt shapes, converts each to a GroupShape with GetResultOfSmartArt, throws an exception if the result is null, optionally repositions the validated GroupShape, and saves the workbook.
// Keywords: Aspose.Cells | .NET | SmartArt conversion | GroupShape validation | GetResultOfSmartArt null check | Excel shape processing | C# | exception handling | shape repositioning
// Common Searches: Aspose.Cells check GetResultOfSmartArt null | Validate SmartArt conversion before moving shape | C# Aspose.Cells SmartArt to GroupShape error handling | How to ensure SmartArt is converted to GroupShape | Throw exception when SmartArt conversion fails Aspose
// Developer Intent: Guarantee that every SmartArt object is successfully turned into a non‑null GroupShape before any further manipulation.
// Use Cases: Batch‑process workbooks and abort saving if any SmartArt fails to convert. | Reposition validated GroupShape objects after confirming conversion success. | Integrate a safety check into automated Excel report generation pipelines.
// AI Prompts: Create a logger that records the IDs of SmartArt shapes returning null from GetResultOfSmartArt. | Write unit tests for the SmartArt validation routine using Aspose.Cells mock objects. | Suggest a fallback strategy that retains the original SmartArt when conversion returns null.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtValidation
{
    // Loads an Excel file, verifies its existence, scans every worksheet for SmartArt shapes, converts each to a GroupShape with GetResultOfSmartArt, throws an exception if the result is null, optionally repositions the validated GroupShape, and saves the workbook.
    public class SmartArtValidator
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWithSmartArt.xlsx";
            const string outputPath = "OutputValidatedSmartArt.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException($"Input file not found: {inputPath}");
            }

            // Load the workbook that may contain SmartArt shapes
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Convert the SmartArt shape to a GroupShape
                        GroupShape groupShape = shape.GetResultOfSmartArt();

                        // Validate that the conversion returned a non‑null GroupShape
                        if (groupShape == null)
                        {
                            throw new InvalidOperationException(
                                $"SmartArt shape with Id {shape.Id} could not be converted to a GroupShape.");
                        }

                        // Example operation after successful validation:
                        // Move the resulting group to a new location
                        groupShape.Left = 200;
                        groupShape.Top = 100;
                    }
                }
            }

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}
