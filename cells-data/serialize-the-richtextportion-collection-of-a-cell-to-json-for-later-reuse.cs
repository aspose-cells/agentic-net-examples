// Title: Serialize a cell's RichTextPortion collection to JSON and restore it with Aspose.Cells for .NET
// Common Searches: how to export Aspose.Cells RichTextPortion font settings to a JSON file in C# | C# Aspose.Cells deserialize cell rich text formatting from JSON | save Excel cell text and font as JSON using Aspose.Cells | Aspose.Cells round‑trip rich text portion data via JSON | convert RichTextPortion collection to JSON and back with Aspose.Cells .NET
// Tags: Aspose.Cells RichTextPortion JSON round‑trip | serialize cell font properties Aspose.Cells C# | deserialize Excel cell formatting from JSON Aspose | RichTextPortion collection export to JSON | restore cell rich text styling Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace RichTextPortionSerialization
{
    // DTO for Font properties that we want to persist
    // The example loads an Excel workbook, extracts the text and full‑cell font from cell A1, maps these values to DTO objects, and serializes a list of RichTextPortionDto to an indented JSON file. It then reads the JSON, deserializes the list, creates a new workbook, applies the stored text and font attributes (name, size, bold, italic, underline, color) to cell A1, and saves the restored workbook.
    public class FontDto
    {
        public string Name { get; set; } = string.Empty;
        public int Size { get; set; }                     // Font size as integer to match Aspose.Cells Font.Size
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }             // true if any underline is applied
        public string ColorArgb { get; set; } = string.Empty; // stored as ARGB integer string
    }

    // DTO for a single text portion (simplified to whole cell text)
    public class RichTextPortionDto
    {
        public string Text { get; set; } = string.Empty;
        public FontDto Font { get; set; } = new FontDto();
    }

    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string jsonPath = "richtext.json";
            const string outputPath = "output.xlsx";

            try
            {
                // -----------------------------------------------------------------
                // Part 1 – Serialize the cell text and its font to JSON
                // -----------------------------------------------------------------
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook that contains text in cell A1
                Workbook srcWorkbook = new Workbook(inputPath);
                Worksheet srcSheet = srcWorkbook.Worksheets[0];
                Cell srcCell = srcSheet.Cells["A1"];

                // Get cell text
                string cellText = srcCell.StringValue ?? string.Empty;

                // Get cell font (applies to the whole cell)
                Font srcFont = srcCell.GetStyle().Font;

                FontDto fontDto = new FontDto
                {
                    Name = srcFont.Name,
                    Size = (int)srcFont.Size, // Cast to int for compatibility
                    IsBold = srcFont.IsBold,
                    IsItalic = srcFont.IsItalic,
                    // Font.Underline is an enum; treat any non‑None value as underline
                    IsUnderline = srcFont.Underline != FontUnderlineType.None,
                    ColorArgb = srcFont.Color.ToArgb().ToString()
                };

                RichTextPortionDto portionDto = new RichTextPortionDto
                {
                    Text = cellText,
                    Font = fontDto
                };

                // Serialize to JSON (indented for readability)
                JsonSerializerOptions jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(new List<RichTextPortionDto> { portionDto }, jsonOptions);
                File.WriteAllText(jsonPath, json);
                Console.WriteLine($"Cell text and font serialized to '{jsonPath}'.");
                Console.WriteLine(json);

                // -----------------------------------------------------------------
                // Part 2 – Deserialize the JSON and apply it to a new workbook
                // -----------------------------------------------------------------
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"JSON file '{jsonPath}' not found.");
                    return;
                }

                string jsonFromFile = File.ReadAllText(jsonPath);
                List<RichTextPortionDto>? deserializedPortions = JsonSerializer.Deserialize<List<RichTextPortionDto>>(jsonFromFile);

                if (deserializedPortions == null || deserializedPortions.Count == 0)
                {
                    Console.WriteLine("No data found in JSON.");
                    return;
                }

                // Create a new workbook and target cell
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];
                Cell destCell = destSheet.Cells["A1"];

                // Apply the deserialized text and font
                RichTextPortionDto firstPortion = deserializedPortions[0];
                destCell.PutValue(firstPortion.Text);

                Style destStyle = destCell.GetStyle();
                Font destFont = destStyle.Font;

                destFont.Name = firstPortion.Font.Name;
                destFont.Size = firstPortion.Font.Size; // Size is now int, matching Font.Size
                destFont.IsBold = firstPortion.Font.IsBold;
                destFont.IsItalic = firstPortion.Font.IsItalic;
                // Restore underline based on the boolean flag
                destFont.Underline = firstPortion.Font.IsUnderline ? FontUnderlineType.Single : FontUnderlineType.None;

                if (int.TryParse(firstPortion.Font.ColorArgb, out int argb))
                {
                    destFont.Color = Color.FromArgb(argb);
                }

                destCell.SetStyle(destStyle);

                // Save the workbook with restored text and font
                destWorkbook.Save(outputPath);
                Console.WriteLine($"Cell text and font restored and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
