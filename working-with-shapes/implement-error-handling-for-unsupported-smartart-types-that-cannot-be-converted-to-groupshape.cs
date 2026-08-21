// Title: Handle Unsupported SmartArt Types When Converting to GroupShape with Aspose.Cells for .NET
// Description: The example loads an Excel workbook, verifies the file, iterates through each worksheet's shapes, and attempts to convert SmartArt objects to GroupShape via GetResultOfSmartArt. It wraps the conversion in try‑catch blocks, logs shape IDs and error messages for unsupported types or null results, and saves the workbook using OoxmlSaveOptions with UpdateSmartArt enabled.
// Keywords: Aspose.Cells | SmartArt | GetResultOfSmartArt | GroupShape | error handling | unsupported SmartArt | UpdateSmartArt | .NET | C# | Excel shape conversion
// Common Searches: Aspose.Cells catch exception GetResultOfSmartArt | convert SmartArt to GroupShape .NET | handle unsupported SmartArt type Aspose.Cells | null GroupShape result check | save workbook with UpdateSmartArt option
// Developer Intent: Add robust try‑catch logic to safely process SmartArt shapes and continue when conversion fails.
// Use Cases: Log the shape Id and error details for each SmartArt conversion that throws an exception. | Skip further manipulation when GetResultOfSmartArt returns null to prevent NullReferenceException. | Persist only successful conversions by saving the workbook with UpdateSmartArt enabled. | Validate the existence of the input file before starting the conversion loop.
// AI Prompts: Write C# code that wraps GetResultOfSmartArt in a custom SmartArtConversionException and logs details to a file. | Refactor the conversion loop using LINQ to filter out unsupported SmartArt types before calling GetResultOfSmartArt. | Explain how OoxmlSaveOptions.UpdateSmartArt influences the saved file when some SmartArt conversions fail. | Generate a PowerShell script that runs the compiled example, captures console output, and writes a summary report.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, verifies the file, iterates through each worksheet's shapes, and attempts to convert SmartArt objects to GroupShape via GetResultOfSmartArt. It wraps the conversion in try‑catch blocks, logs shape IDs and error messages for unsupported types or null results, and saves the workbook using OoxmlSaveOptions with UpdateSmartArt enabled.
    public class SmartArtConversionWithErrorHandling
    {
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
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that may contain SmartArt shapes
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Iterate through all worksheets and their shapes
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        try
                        {
                            // Convert the SmartArt to a GroupShape
                            GroupShape groupShape = shape.GetResultOfSmartArt();

                            if (groupShape != null)
                            {
                                // Example modification: move the resulting group shape
                                groupShape.Left += 50;
                                groupShape.Top += 20;
                            }
                            else
                            {
                                Console.WriteLine($"SmartArt shape (Id={shape.Id}) could not be converted (null result).");
                            }
                        }
                        catch (Exception ex)
                        {
                            // Handle unsupported SmartArt types or conversion errors
                            Console.WriteLine($"Error converting SmartArt shape (Id={shape.Id}): {ex.Message}");
                        }
                    }
                }
            }

            // Save the workbook with UpdateSmartArt enabled to persist conversions
            try
            {
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true
                };
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
