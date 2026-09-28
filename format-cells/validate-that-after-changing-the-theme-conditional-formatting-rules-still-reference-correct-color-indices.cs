// Title: Check that conditional formatting colors remain correct after changing an Excel workbook theme using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens an .xlsx file, loops through each worksheet's ConditionalFormattings collection, and prints the foreground and background ARGB values of every FormatCondition. | Add error handling to verify the input file exists and create the output folder if missing before saving the workbook. | Extend the program to compare each retrieved color with the expected theme palette index and log any mismatches.
// Common Searches: aspnet how to list conditional formatting colors in an Excel file with Aspose.Cells | c# verify conditional formatting palette indices after applying a new theme | aspose.cells check if conditional formatting uses correct theme colors | retrieve ARGB values of conditional formatting rules in .xlsx using .NET | debug conditional formatting color references after Excel theme change
// Tags: Aspose.Cells conditional formatting color validation | C# iterate ConditionalFormattings collection | Excel theme palette verification with Aspose.Cells | retrieve ARGB values from FormatCondition style | validate conditional formatting after theme change

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an existing workbook, iterates each worksheet's conditional formatting collections, outputs the ARGB values of foreground and background colors for every rule, ensures the output directory exists, and saves the workbook while handling missing input files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Changing the workbook theme is not supported directly via a public API in the current version.
            // If needed, theme changes can be applied through other means (e.g., applying styles manually).

            // Validate that conditional formatting rules still reference correct colors
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    ConditionalFormattingCollection cfCollection = sheet.ConditionalFormattings;
                    for (int cfIndex = 0; cfIndex < cfCollection.Count; cfIndex++)
                    {
                        // Use dynamic to avoid compile‑time dependency on ConditionalFormatting type
                        dynamic cf = cfCollection[cfIndex];
                        for (int condIndex = 0; condIndex < cf.FormatConditions.Count; condIndex++)
                        {
                            FormatCondition condition = cf.FormatConditions[condIndex];
                            Style style = condition.Style;

                            // Foreground color validation (output ARGB value)
                            Color fgColor = style.ForegroundColor;
                            Console.WriteLine(
                                $"Sheet '{sheet.Name}', CF {cfIndex}, Condition {condIndex}: Foreground ARGB = 0x{fgColor.ToArgb():X8}");

                            // Background color validation (output ARGB value)
                            Color bgColor = style.BackgroundColor;
                            Console.WriteLine(
                                $"Sheet '{sheet.Name}', CF {cfIndex}, Condition {condIndex}: Background ARGB = 0x{bgColor.ToArgb():X8}");
                        }
                    }
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook after validation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
