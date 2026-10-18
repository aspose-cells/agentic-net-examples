// Title: Set header row font color to the workbook theme Accent1 across all worksheets using Aspose.Cells for .NET
// Common Searches: how to change the first row font color to the workbook theme Accent1 in every sheet with Aspose.Cells C# | apply workbook theme Accent1 color to header rows of multiple worksheets using Aspose.Cells | C# iterate worksheets and set header row font color from theme with Aspose.Cells | retrieve theme Accent1 color and style Excel header row in Aspose.Cells .NET | range.ApplyStyle FontColor StyleFlag example for header rows in Aspose.Cells
// Tags: apply theme accent1 font color to Excel header row | iterate all worksheets and style first row Aspose.Cells | retrieve workbook theme color using GetThemeColor | range.ApplyStyle with FontColor StyleFlag in C# | set header row font color based on workbook theme

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For ThemeColorType

namespace AsposeCellsExample
{
    // The example loads an existing XLSX workbook, obtains the Accent1 color from the workbook's theme (fallback to black), creates a style that uses this color for the font, iterates through every worksheet, defines a range covering the first row, and applies the style with a StyleFlag that targets only the font color before saving the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the Accent1 color from the workbook's theme (fallback to Black if unavailable)
                Color accent1Color = Color.Black;
                try
                {
                    // Aspose.Cells provides GetThemeColor to obtain theme colors
                    accent1Color = workbook.GetThemeColor(ThemeColorType.Accent1);
                }
                catch
                {
                    // If any theme‑related API is not supported, keep the fallback color
                }

                // Create a style that uses the Accent1 color for the font
                Style accentStyle = workbook.CreateStyle();
                accentStyle.Font.Color = accent1Color;

                // Specify that only the font color should be applied
                StyleFlag styleFlag = new StyleFlag
                {
                    FontColor = true
                };

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Determine the last used column index in the sheet
                        int lastColumn = sheet.Cells.MaxColumn;

                        // If the sheet is empty, skip styling
                        if (lastColumn < 0)
                            continue;

                        // Create a range that covers the entire header row (row 0)
                        Aspose.Cells.Range headerRange = sheet.Cells.CreateRange(0, 0, 1, lastColumn + 1);

                        // Apply the style to the header row
                        headerRange.ApplyStyle(accentStyle, styleFlag);
                    }
                    catch (Exception exSheet)
                    {
                        Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook to the output file
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
