// Title: How to deserialize a JSON file with rich‑text segments and reapply the formatting to an Excel cell using Aspose.Cells for .NET
// Common Searches: C# deserialize rich text JSON and set character formatting in an Aspose.Cells worksheet | restore cell rich‑text formatting from a JSON export with Aspose.Cells .NET | apply font properties to a range of characters in an Excel cell using the Aspose.Cells API | read saved rich‑text segment data from JSON and reapply to cell A1 in C# | Aspose.Cells example for reloading rich‑text formatting after exporting to JSON
// Tags: Aspose.Cells deserialize rich‑text JSON | C# character‑level font formatting in Excel | reapply cell rich‑text using FontSetting | Excel cell rich‑text restoration .NET | Aspose.Cells FontInfo mapping from JSON

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace RichTextFromJsonDemo
{
    // Model that matches the JSON structure exported by Aspose.Cells for a cell's rich‑text.
    // The sample loads an existing workbook, reads a JSON file that contains a list of RichTextSegment objects exported by Aspose.Cells, deserializes it into C# models, and iterates over each segment. For each segment it obtains the character range with Cell.Characters(start, length) and applies the stored font name, size, color, bold, italic, underline, and strikeout settings via the FontSetting.Font object. Finally, the workbook is saved with the restored rich‑text formatting.
    public class RichTextSegment
    {
        public int Start { get; set; }               // Zero‑based start index of the segment
        public int Length { get; set; }              // Number of characters in the segment
        public FontInfo? Font { get; set; }          // Font information for this segment (nullable)
    }

    public class FontInfo
    {
        public string? Name { get; set; }
        public double Size { get; set; }
        public int ColorArgb { get; set; }           // ARGB integer (e.g., 0xFF0000FF for blue)
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public bool IsStrikeout { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            const string workbookPath = "input.xlsx";
            const string jsonPath = "cellRichText.json";
            const string outputPath = "output.xlsx";

            try
            {
                // -----------------------------------------------------------------
                // 1. Load the workbook that contains the target cell.
                // -----------------------------------------------------------------
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                Workbook workbook = new Workbook(workbookPath);
                Worksheet worksheet = workbook.Worksheets[0]; // assuming first sheet
                Cell targetCell = worksheet.Cells["A1"];      // change address as needed

                // -----------------------------------------------------------------
                // 2. Read and deserialize the JSON that holds rich‑text information.
                // -----------------------------------------------------------------
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"JSON file not found: {jsonPath}");
                    return;
                }

                string jsonContent = File.ReadAllText(jsonPath);
                List<RichTextSegment>? segments = JsonSerializer.Deserialize<List<RichTextSegment>>(jsonContent,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                if (segments == null || segments.Count == 0)
                {
                    Console.WriteLine("No rich‑text segments were found in the JSON.");
                    return;
                }

                // -----------------------------------------------------------------
                // 3. Apply each segment's font settings back to the cell.
                // -----------------------------------------------------------------
                foreach (var seg in segments)
                {
                    // Validate indices
                    if (seg.Start < 0 || seg.Length <= 0)
                    {
                        Console.WriteLine($"Skipping invalid segment (Start={seg.Start}, Length={seg.Length}).");
                        continue;
                    }

                    try
                    {
                        // Obtain the FontSetting object for the specified range.
                        FontSetting fontSetting = targetCell.Characters(seg.Start, seg.Length);
                        Font font = fontSetting.Font;

                        if (seg.Font != null)
                        {
                            // Apply font name
                            if (!string.IsNullOrEmpty(seg.Font.Name))
                                font.Name = seg.Font.Name;

                            // Apply font size (Aspose.Cells.Font.Size is an integer)
                            font.Size = (int)seg.Font.Size;

                            // Apply font color
                            font.Color = Color.FromArgb(seg.Font.ColorArgb);

                            // Apply style flags
                            font.IsBold = seg.Font.IsBold;
                            font.IsItalic = seg.Font.IsItalic;
                            font.IsStrikeout = seg.Font.IsStrikeout;

                            // Apply underline (Aspose.Cells uses FontUnderlineType)
                            font.Underline = seg.Font.IsUnderline
                                ? FontUnderlineType.Single
                                : FontUnderlineType.None;
                        }
                    }
                    catch (Exception segEx)
                    {
                        Console.WriteLine($"Error applying segment starting at {seg.Start}: {segEx.Message}");
                    }
                }

                // -----------------------------------------------------------------
                // 4. Save the workbook with the restored rich‑text.
                // -----------------------------------------------------------------
                workbook.Save(outputPath);
                Console.WriteLine($"Rich‑text successfully reapplied and workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
