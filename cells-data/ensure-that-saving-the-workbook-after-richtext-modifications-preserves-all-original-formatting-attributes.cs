// Title: How to modify cell A1 rich‑text and save the workbook while preserving all original formatting with Aspose.Cells for .NET
// AI Prompts: Replace the first two words in cell A1 with bold red and italic blue styling using FontSetting, then save the workbook so that every existing rich‑text attribute stays unchanged. | Apply character‑level font color and style changes to a cell’s text and export the Excel file without losing any prior formatting applied by Aspose.Cells.
// Common Searches: Aspose.Cells .NET keep existing rich text formatting after updating cell value | preserve character level font settings when saving modified Excel workbook using Aspose | how to save workbook without stripping rich‑text attributes in C# Aspose.Cells | update specific words in a cell with different fonts and retain other formatting Aspose.Cells | C# Aspose.Cells save workbook preserving original cell styles and rich text
// Tags: modify cell rich text Aspose.Cells | character level formatting Excel .NET | preserve original formatting on workbook save | fontsetting usage Aspose.Cells | save workbook without losing rich text

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRichTextExample
{
    // The example loads an existing Excel file, extracts the first two words from cell A1, reassembles them, applies bold red formatting to the first word and italic blue to the second using FontSetting, and then saves the workbook while ensuring all other original rich‑text and cell formatting remain intact.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input workbook exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook while preserving all original formatting.
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A1"];

                // Preserve any existing text.
                string currentText = cell.StringValue ?? string.Empty;
                string[] parts = currentText.Split(' ');

                if (parts.Length >= 2)
                {
                    // Re‑assemble the cell value (first two words separated by a space).
                    string newText = $"{parts[0]} {parts[1]}";
                    cell.PutValue(newText);

                    // Retrieve character‑level font settings.
                    FontSetting[] charSettings = cell.GetCharacters();

                    // Apply formatting to the first part (bold red).
                    int part1Length = parts[0].Length;
                    for (int i = 0; i < part1Length && i < charSettings.Length; i++)
                    {
                        charSettings[i].Font.Color = Color.Red;
                        charSettings[i].Font.IsBold = true;
                    }

                    // Apply formatting to the second part (italic blue).
                    int part2Start = part1Length + 1; // skip the space
                    int part2Length = parts[1].Length;
                    for (int i = part2Start; i < part2Start + part2Length && i < charSettings.Length; i++)
                    {
                        charSettings[i].Font.Color = Color.Blue;
                        charSettings[i].Font.IsItalic = true;
                    }
                }
                else
                {
                    // No sufficient parts – keep original text unchanged.
                    cell.PutValue(currentText);
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
