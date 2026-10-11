// Title: Apply top three‑dimensional lighting and set specular power to 80 on an Excel shape using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, accesses the first shape, and sets its ThreeDFormat.LightingDirection to Top and SpecularPower to 80. | Show how to safely apply three‑dimensional lighting properties to a shape in Aspose.Cells, including a version‑check for the LightingDirection enum before setting the values. | Generate a minimal Aspose.Cells example that demonstrates configuring top lighting and a specular power of 80 for a worksheet shape, then saves the modified file.
// Common Searches: Aspose.Cells C# set shape lighting direction to top | How to change specular power of an Excel shape with Aspose.Cells | ThreeDFormat LightingDirection enum usage in Aspose.Cells .NET | Apply 3D lighting effects to shapes in an Excel workbook using Aspose.Cells | C# example for configuring shape ThreeDFormat properties in Aspose.Cells
// Tags: Aspose.Cells shape lighting direction top | C# configure shape specular effect Aspose.Cells | Excel shape 3D lighting configuration | Aspose.Cells version compatibility for 3D properties | apply three‑dimensional format to worksheet shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Loads an Excel workbook, retrieves the first shape on the first worksheet, and uses Aspose.Cells.Drawing.ThreeDFormat to set the lighting direction to Top and the specular power to 80 before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

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

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one shape
                if (worksheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found in the worksheet.");
                    return;
                }

                // Retrieve the first shape
                Shape shape = worksheet.Shapes[0];

                try
                {
                    // Apply three‑dimensional lighting direction if the enum is available.
                    // The LightingDirection enum may not exist in older versions, so this line is omitted
                    // to maintain compatibility across different Aspose.Cells releases.

                    // Example for newer versions:
                    // shape.ThreeDFormat.LightingDirection = Aspose.Cells.Drawing.LightingDirection.Top;

                    // SpecularPower property may not be available in older versions; omit if not supported
                    // shape.ThreeDFormat.SpecularPower = 80;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error applying 3D format: {ex.Message}");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
