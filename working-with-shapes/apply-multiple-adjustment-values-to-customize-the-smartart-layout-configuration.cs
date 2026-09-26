// Title: Detect SmartArt shapes in an Excel worksheet and handle unsupported adjustment operations with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that opens an existing .xlsx file, retrieves a shape by index, determines if it is a SmartArt object, and logs a notice when SmartArt adjustment APIs are unavailable. | Create a robust C# routine that validates the input workbook path, checks that the requested shape index is within the worksheet's shape collection, identifies SmartArt shapes, and safely saves the workbook after processing. | Write a C# method that returns true when a given shape in an Aspose.Cells worksheet is a SmartArt element and throws a custom exception if adjustment functionality is not supported.
// Common Searches: C# Aspose.Cells how to check if a shape is SmartArt in an Excel file | detect SmartArt objects in worksheet using Aspose.Cells .NET | Aspose.Cells SmartArt adjustment API not available workaround | validate shape index before accessing shapes collection Aspose.Cells | save Excel workbook after processing shapes with Aspose.Cells
// Tags: Aspose.Cells detect SmartArt shape | C# validate shape index worksheet | SmartArt adjustment unsupported Aspose.Cells | Excel workbook load save Aspose.Cells | shape type inspection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtAdjustment
{
    // The example loads an existing Excel workbook with Aspose.Cells, accesses the first worksheet, retrieves a shape by a specified index, checks the shape's type name for "SmartArt", logs that adjustment operations are not supported in the current API version, ensures the output directory exists, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = @"C:\Temp\InputWorkbook.xlsx";
            string outputPath = @"C:\Temp\OutputWorkbook.xlsx";

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
                Worksheet worksheet = workbook.Worksheets[0];

                // Index of the SmartArt shape within the worksheet's shape collection
                int smartArtShapeIndex = 0; // Change this to target the correct shape

                // Ensure the index is within bounds
                if (smartArtShapeIndex < 0 || smartArtShapeIndex >= worksheet.Shapes.Count)
                {
                    Console.WriteLine("SmartArt shape index is out of range.");
                    return;
                }

                // Retrieve the shape
                Shape shape = worksheet.Shapes[smartArtShapeIndex];

                // Determine whether the shape is a SmartArt object.
                // Aspose.Cells does not expose a dedicated ShapeType for SmartArt in some versions,
                // so we inspect the type name as a fallback.
                if (shape.Type.ToString().IndexOf("SmartArt", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // NOTE: Aspose.Cells does not expose direct adjustment APIs for SmartArt.
                    // If adjustment functionality becomes available in future versions,
                    // it can be applied here. For now we simply acknowledge the shape type.
                    Console.WriteLine("SmartArt shape detected. Adjustments are not supported in this API version.");
                }
                else
                {
                    Console.WriteLine("The specified shape is not a SmartArt object.");
                }

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the (potentially) modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
