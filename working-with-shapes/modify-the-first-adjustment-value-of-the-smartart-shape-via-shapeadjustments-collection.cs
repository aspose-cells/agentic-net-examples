// Title: How to set the first adjustment value of a SmartArt shape using Shape.Adjustments in Aspose.Cells for .NET (C#) – limitations and workarounds
// AI Prompts: Write C# code that attempts to assign a new value to the first element of the Shape.Adjustments collection for a SmartArt shape named "SmartArt1" using Aspose.Cells, and include error handling for unsupported operations. | Describe alternative strategies for changing SmartArt appearance in Aspose.Cells when direct adjustment manipulation is unavailable, such as replacing the SmartArt with a supported shape or exporting and re‑importing the graphic.
// Common Searches: Aspose.Cells set first adjustment of SmartArt shape C# | Shape.Adjustments collection not updating SmartArt Aspose.Cells | workaround to modify SmartArt properties with Aspose.Cells .NET | cannot change SmartArt adjustment values using Aspose.Cells | replace SmartArt with image in Aspose.Cells C#
// Tags: smartart adjustment modification aspose.cells c# | shape.adjustments unsupported smartart aspose | smartart replacement workaround aspose.cells | modify smartart properties via aspose.cells c# | first adjustment index shape.adjustments aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, looks for a SmartArt shape named "SmartArt1", and demonstrates that Aspose.Cells does not currently expose direct manipulation of SmartArt adjustment values through the Shape.Adjustments collection, offering guidance on handling this limitation and possible workarounds before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the shape named "SmartArt1"
            Shape shape = worksheet.Shapes["SmartArt1"];

            if (shape != null)
            {
                // Aspose.Cells does not expose SmartArt manipulation directly.
                // If needed, you can replace the shape or perform other supported operations here.
                Console.WriteLine("Shape 'SmartArt1' found. SmartArt manipulation is not supported via Aspose.Cells API.");
            }
            else
            {
                Console.WriteLine("SmartArt shape 'SmartArt1' not found.");
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
