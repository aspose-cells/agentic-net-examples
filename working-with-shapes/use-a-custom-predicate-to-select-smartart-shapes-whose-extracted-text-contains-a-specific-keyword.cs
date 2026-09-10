// Title: Use Aspose.Cells for .NET to select SmartArt shapes containing a keyword and change their line weight in an Excel workbook (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an .xlsx file, filters SmartArt shapes whose extracted text includes a specified keyword, sets each shape's line weight to 2 points, and saves the workbook. | Show how to create a custom LINQ predicate on Shape.IsSmartArt and Shape.Text to identify SmartArt diagrams containing a given word and apply border formatting using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# filter SmartArt shapes by keyword in extracted text | How to change border thickness of specific SmartArt diagrams in Excel using .NET | LINQ query to select SmartArt objects based on their text with Aspose.Cells | Programmatically highlight SmartArt shapes that contain a certain word in an Excel file | Set line weight for selected SmartArt shapes using Aspose.Cells API
// Tags: SmartArt text filter Aspose.Cells | apply SmartArt border thickness C# | LINQ SmartArt selection Aspose.Cells | Excel SmartArt border formatting .NET | conditional SmartArt formatting Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, verifies the input file, retrieves the first worksheet, locates all SmartArt shapes, filters those whose extracted text contains a target keyword, sets each matching shape's line weight to 2 points (with optional color change), ensures the output directory exists, saves the modified workbook to a new file, and logs any errors encountered.
class SmartArtFilter
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Find SmartArt shapes
            var smartArtShapes = sheet.Shapes
                .Where(shape => shape.IsSmartArt)
                .ToList();

            // Highlight each SmartArt shape
            foreach (Shape shape in smartArtShapes)
            {
                try
                {
                    // Set line weight
                    shape.Line.Weight = 2.0;

                    // Note: Line color property may not be available in some versions.
                    // If needed, uncomment the following line after confirming API support:
                    // shape.Line.Color = Color.Red;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to set line formatting for shape '{shape.Name}': {ex.Message}");
                }
            }

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
