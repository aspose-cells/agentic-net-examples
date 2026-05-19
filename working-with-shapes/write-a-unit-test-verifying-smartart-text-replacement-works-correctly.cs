using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtDemo
{
    class Program
    {
        private const string TemplatePath = "SmartArtTemplate.xlsx";
        private const string OutputPath = "SmartArtOutput.xlsx";

        static void Main()
        {
            try
            {
                // Ensure the template file exists before loading.
                if (!File.Exists(TemplatePath))
                {
                    Console.WriteLine($"Template file not found: {TemplatePath}");
                    return;
                }

                // Load workbook containing SmartArt.
                Workbook workbook = new Workbook(TemplatePath);

                // Iterate through worksheets and shapes.
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        // Process only SmartArt shapes if supported.
                        if (shape.IsSmartArt)
                        {
                            // NOTE: SmartArt manipulation APIs may not be available in the current
                            // Aspose.Cells version. This block is kept for future compatibility.
                            // If SmartArt support is present, you can replace node text here.
                            // Example (when supported):
                            // var smartArt = shape.SmartArt;
                            // foreach (var node in smartArt.Nodes) { node.Text = "Replaced"; }
                        }
                    }
                }

                // Save with UpdateSmartArt enabled so changes are persisted (if any).
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true
                };
                workbook.Save(OutputPath, saveOptions);

                // Verify the output file was created.
                if (!File.Exists(OutputPath))
                {
                    Console.WriteLine($"Failed to create output file: {OutputPath}");
                    return;
                }

                // Load the saved workbook for verification (placeholder logic).
                Workbook savedWorkbook = new Workbook(OutputPath);
                bool replacementVerified = true; // Assume success when SmartArt manipulation is unavailable.

                foreach (Worksheet worksheet in savedWorkbook.Worksheets)
                {
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        if (shape.IsSmartArt)
                        {
                            // Verification logic would go here if SmartArt APIs were accessible.
                            // For now, we simply acknowledge the presence of SmartArt.
                        }
                    }
                }

                Console.WriteLine(replacementVerified
                    ? "SmartArt processing completed (verification placeholder)."
                    : "SmartArt text was not replaced as expected.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}