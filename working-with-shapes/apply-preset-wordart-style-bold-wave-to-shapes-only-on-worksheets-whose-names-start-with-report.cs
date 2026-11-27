// Title: Apply the Bold Wave WordArt preset to shapes on Excel worksheets whose names start with "Report" using Aspose.Cells for .NET
// Common Searches: how to apply Bold Wave WordArt to shapes in specific worksheets using Aspose.Cells C# | filter worksheets by name prefix and set WordArt style on AutoShape with Aspose.Cells | C# code to set PresetWordArtStyle.WordArtStyle2 for shapes on Report* sheets | Aspose.Cells example for applying WordArt style only to shapes with TextBody | apply WordArt preset to Excel shapes conditionally based on worksheet name in .NET
// Tags: apply preset WordArt Bold Wave to Excel shapes | conditional worksheet name filtering Aspose.Cells | set WordArt style via SetWordArtStyle method | shape TextBody validation before styling | process only worksheets starting with Report in .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace ApplyBoldWaveWordArt
{
    // The sample loads an Excel workbook, iterates through each worksheet, and processes only those whose names begin with "Report" (case‑insensitive). For each qualifying sheet it loops over all shapes, checks that the shape has a TextBody, and applies the preset WordArt style "Bold Wave" using SetWordArtStyle(PresetWordArtStyle.WordArtStyle2). The modified workbook is then saved to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths – adjust as needed.
            const string inputFile = "ReportWorkbook.xlsx";
            const string outputFile = "ReportWorkbook_Styled.xlsx";

            // Load the workbook.
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputFile);
                Console.WriteLine($"Workbook \"{inputFile}\" loaded successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook \"{inputFile}\": {ex.Message}");
                return;
            }

            // Iterate through all worksheets.
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Process only worksheets whose name starts with "Report" (case‑insensitive).
                if (!worksheet.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
                    continue;

                Console.WriteLine($"Applying Bold Wave WordArt style to shapes on worksheet \"{worksheet.Name}\".");

                // Iterate through all shapes on the worksheet.
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Only shapes that support a TextBody can receive a WordArt style.
                    // TextBox, AutoShape, and existing WordArt objects expose TextBody.
                    if (shape.TextBody == null)
                    {
                        Console.WriteLine($"  Skipping shape \"{shape.Name}\" (no TextBody).");
                        continue;
                    }

                    try
                    {
                        // Apply the preset WordArt style "Bold Wave".
                        // The exact enum value for "Bold Wave" is PresetWordArtStyle.WordArtStyle2
                        // (based on Aspose.Cells documentation). If a different style is required,
                        // replace the enum member accordingly.
                        shape.TextBody.SetWordArtStyle(PresetWordArtStyle.WordArtStyle2);
                        Console.WriteLine($"  Shape \"{shape.Name}\" styled successfully.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Failed to style shape \"{shape.Name}\": {ex.Message}");
                    }
                }
            }

            // Save the modified workbook.
            try
            {
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved as \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook \"{outputFile}\": {ex.Message}");
            }
        }
    }
}
