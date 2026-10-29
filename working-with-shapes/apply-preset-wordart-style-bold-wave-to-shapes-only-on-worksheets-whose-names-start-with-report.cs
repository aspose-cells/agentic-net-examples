// Title: Apply a bold text effect to shapes on Excel worksheets whose names start with "Report" using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape text to bold on worksheets with name starting Report | Iterate Excel worksheets and apply TextEffect.FontBold to shapes only on report sheets using .NET | How to change font size of shapes in specific worksheets with Aspose.Cells | Apply text formatting to shapes in Excel files based on worksheet name prefix in C# | Conditional formatting of shape TextEffect properties for worksheets that begin with Report
// Tags: shape text effect modification Aspose.Cells | worksheet name prefix filtering Aspose.Cells | set shape FontBold property C# | adjust shape TextEffect font size Aspose.Cells | conditional shape styling Excel .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, filters worksheets whose names start with "Report", and for each shape on those sheets sets the TextEffect.FontBold flag to true and FontSize to 14. Shape‑specific errors are caught so processing continues, and the modified workbook is saved to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Process only worksheets whose names start with "Report"
                    if (sheet.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
                    {
                        // Iterate through all shapes on the worksheet
                        foreach (Shape shape in sheet.Shapes)
                        {
                            try
                            {
                                // Apply a simple text effect (bold) to the shape as a safe alternative to WordArt
                                shape.TextEffect.FontBold = true;
                                shape.TextEffect.FontSize = 14;
                            }
                            catch (Exception shapeEx)
                            {
                                // Log shape‑specific errors but continue processing other shapes
                                Console.Error.WriteLine($"Failed to modify shape '{shape.Name}': {shapeEx.Message}");
                            }
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
