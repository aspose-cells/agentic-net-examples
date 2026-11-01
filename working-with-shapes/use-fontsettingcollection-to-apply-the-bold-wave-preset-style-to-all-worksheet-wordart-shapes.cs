// Title: Apply bold and italic formatting to every WordArt shape across all worksheets with Aspose.Cells for .NET
// Common Searches: how to set bold and italic on WordArt shapes in an Excel file using Aspose.Cells C# | iterate through workbook shapes and modify TextEffect properties with Aspose.Cells | programmatically change WordArt font style for all sheets in .NET | Aspose.Cells C# example to format WordArt text across multiple worksheets | apply text effect formatting to Excel WordArt using Aspose.Cells API
// Tags: WordArt TextEffect formatting Aspose.Cells | apply bold italic to Excel shapes C# | iterate workbook shapes Aspose.Cells .NET | modify WordArt font properties programmatically | Aspose.Cells shape styling across worksheets

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, loops through each worksheet and its shapes, detects WordArt shapes via the TextEffect property, sets FontBold and FontItalic to true, and saves the updated file, handling missing input files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Identify WordArt shapes by checking the TextEffect property
                        if (shape.TextEffect != null)
                        {
                            // Apply simple formatting to the WordArt text
                            // (Using properties that are available across Aspose.Cells versions)
                            shape.TextEffect.FontBold = true;
                            shape.TextEffect.FontItalic = true;
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
