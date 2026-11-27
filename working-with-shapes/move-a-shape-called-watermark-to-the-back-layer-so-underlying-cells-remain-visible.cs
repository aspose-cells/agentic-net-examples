// Title: Move a Watermark shape to the back layer in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells move watermark shape behind cells | How to send a shape to the back layer in an Excel file using Aspose.Cells | Aspose.Cells ToFrontOrBack example for shape layering | Set shape Z-order to back in .xlsx with Aspose.Cells .NET | Remove watermark overlay by moving shape to back in Aspose.Cells workbook
// Tags: move shape to back Aspose.Cells | shape Z-order manipulation C# | Aspose.Cells ToFrontOrBack method | Excel watermark shape layering | adjust shape order in XLSX workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace MoveWatermarkToBack
{
    // The example loads an XLSX file with Aspose.Cells, iterates through each worksheet's shapes, finds the shape named "Watermark", uses shape.ToFrontOrBack(-1000) to send it to the back layer so underlying cells stay visible, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook '{inputPath}': {ex.Message}");
                return;
            }

            bool watermarkFound = false;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // The Shape class has a Name property (inherited from the base class)
                    // Compare with the expected watermark name
                    if (string.Equals(shape.Name, "Watermark", StringComparison.OrdinalIgnoreCase))
                    {
                        // Send the shape to the back layer.
                        // Using a large negative number ensures it goes to the very back.
                        shape.ToFrontOrBack(-1000);
                        Console.WriteLine($"Watermark shape moved to back on worksheet '{sheet.Name}'.");
                        watermarkFound = true;
                        // If there could be multiple watermarks, continue searching;
                        // otherwise break out of loops.
                    }
                }
            }

            if (!watermarkFound)
            {
                Console.WriteLine("No shape named 'Watermark' was found in the workbook.");
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
