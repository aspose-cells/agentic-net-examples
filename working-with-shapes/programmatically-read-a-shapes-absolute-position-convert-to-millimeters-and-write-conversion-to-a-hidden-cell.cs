// Title: Read a shape's absolute position in points, convert to millimeters, and write the values to hidden cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get shape left and top coordinates | convert shape position from points to millimeters in Aspose.Cells | write shape coordinates to a hidden column in Excel with Aspose.Cells | hide Excel column programmatically after inserting data using Aspose.Cells | retrieve shape by name with fallback to first shape Aspose.Cells C#
// Tags: shape position points to millimeters Aspose.Cells | write shape coordinates to hidden Excel column | retrieve shape by name Aspose.Cells C# | hide column programmatically Aspose.Cells | points to millimeters conversion C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, locates a shape by its name (or uses the first shape if not found), reads the shape's Left and Top properties expressed in points, converts those coordinates to millimeters, writes the X and Y values to hidden cells Z1 and Z2, hides column Z, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input and output file paths
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

                // Retrieve the shape (by name or first shape as fallback)
                string shapeName = "MyShape";
                Shape shape = null;

                // Search for shape by name
                foreach (Shape s in worksheet.Shapes)
                {
                    if (s.Name == shapeName)
                    {
                        shape = s;
                        break;
                    }
                }

                // Fallback to first shape if not found
                if (shape == null && worksheet.Shapes.Count > 0)
                {
                    shape = worksheet.Shapes[0];
                    Console.WriteLine($"Shape \"{shapeName}\" not found. Using first shape \"{shape.Name}\".");
                }

                if (shape == null)
                {
                    Console.WriteLine("No shapes found in the worksheet.");
                    return;
                }

                // Get shape position in points (1 point = 1/72 inch)
                double positionXPoints;
                double positionYPoints;

                try
                {
                    // Aspose.Cells provides Left and Top properties for position in points
                    positionXPoints = shape.Left;
                    positionYPoints = shape.Top;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to retrieve shape position: {ex.Message}");
                    return;
                }

                // Convert points to millimeters (1 inch = 25.4 mm)
                const double mmPerPoint = 25.4 / 72.0;
                double positionXmm = positionXPoints * mmPerPoint;
                double positionYmm = positionYPoints * mmPerPoint;

                // Write converted values to hidden cells Z1 and Z2
                worksheet.Cells["Z1"].PutValue(positionXmm);
                worksheet.Cells["Z2"].PutValue(positionYmm);

                // Hide column Z (zero‑based index 25)
                worksheet.Cells.HideColumn(25);

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
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
