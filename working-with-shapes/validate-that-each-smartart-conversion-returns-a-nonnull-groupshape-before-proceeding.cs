// Title: Validate SmartArt‑to‑GroupShape conversion with Aspose.Cells for .NET
// Description: C# sample that loads an Excel file, scans every worksheet for SmartArt shapes, converts each shape with GetResultOfSmartArt, verifies the result is not null, optionally repositions or renames the resulting GroupShape, logs any failures, and saves the workbook using OoxmlSaveOptions with UpdateSmartArt enabled.
// Keywords: Aspose.Cells | SmartArt conversion | GroupShape | GetResultOfSmartArt | C# | .NET Excel automation | null check | shape validation | OoxmlSaveOptions | UpdateSmartArt
// Common Searches: Aspose.Cells verify SmartArt conversion is not null | C# GetResultOfSmartArt returns null handling | move GroupShape after SmartArt conversion Aspose | save workbook with updated SmartArt Aspose.Cells | iterate worksheets and validate SmartArt shapes
// Developer Intent: Confirm that every SmartArt object can be transformed into a GroupShape before applying further modifications.
// Use Cases: Batch‑process Excel workbooks to ensure SmartArt shapes are safely converted for downstream editing. | Log indices of shapes that fail conversion to aid debugging of complex spreadsheets. | Reposition or rename GroupShape objects only when the conversion succeeds, preserving workbook integrity.
// AI Prompts: Generate a C# routine that throws a custom exception if GetResultOfSmartArt returns null for any SmartArt shape. | Create code to collect all SmartArt shapes that could not be converted and output a summary report. | Write unit tests in NUnit that assert GetResultOfSmartArt never returns null for a set of sample SmartArt diagrams using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

namespace AsposeCellsSmartArtValidation
{
    // C# sample that loads an Excel file, scans every worksheet for SmartArt shapes, converts each shape with GetResultOfSmartArt, verifies the result is not null, optionally repositions or renames the resulting GroupShape, logs any failures, and saves the workbook using OoxmlSaveOptions with UpdateSmartArt enabled.
    public class SmartArtValidator
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWithSmartArt.xlsx";
            const string outputPath = "OutputValidatedSmartArt.xlsx";

            // Ensure the input file exists before attempting to load
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load an existing workbook that may contain SmartArt shapes
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Convert SmartArt to a GroupShape
                        GroupShape groupShape = shape.GetResultOfSmartArt();

                        // Validate that the conversion succeeded
                        if (groupShape != null)
                        {
                            // Example operation: move the group to a new location
                            groupShape.Left = 200;
                            groupShape.Top = 100;

                            // Example operation: change alternative text
                            groupShape.AlternativeText = "Converted SmartArt Group";
                        }
                        else
                        {
                            // Handle the case where conversion failed (null result)
                            Console.WriteLine($"SmartArt shape at index {sheet.Shapes.IndexOf(shape)} could not be converted.");
                        }
                    }
                }
            }

            try
            {
                // Save the workbook, optionally updating SmartArt in the saved file
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true
                };
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
