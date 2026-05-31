using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsRichTextReapply
{
    // Classes that match the JSON structure produced by Cell.ToJson()
    public class CellJson
    {
        public CharacterJson[] Characters { get; set; }
    }

    public class CharacterJson
    {
        public int StartIndex { get; set; }
        public int Length { get; set; }
        public FontJson Font { get; set; }
    }

    public class FontJson
    {
        public bool? IsBold { get; set; }
        public bool? IsItalic { get; set; }
        public string Color { get; set; } // Expected in HTML hex format, e.g., "#FF0000"
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cell cell = worksheet.Cells["A1"]; // Target cell

                // Ensure the cell contains plain text before applying character formatting.
                cell.PutValue("Sample rich text for reapplication");

                // Path to the JSON file that stores rich‑text information.
                string jsonPath = "cellRichText.json";

                if (File.Exists(jsonPath))
                {
                    // Read and deserialize the JSON.
                    string jsonContent = File.ReadAllText(jsonPath);
                    CellJson cellData = JsonSerializer.Deserialize<CellJson>(jsonContent);

                    // Reapply character formatting if data is present.
                    if (cellData?.Characters != null && cellData.Characters.Length > 0)
                    {
                        FontSetting[] fontSettings = new FontSetting[cellData.Characters.Length];

                        for (int i = 0; i < cellData.Characters.Length; i++)
                        {
                            CharacterJson ch = cellData.Characters[i];

                            // Get a FontSetting for the specified character range.
                            FontSetting setting = cell.Characters(ch.StartIndex, ch.Length);

                            // Apply font properties from JSON.
                            if (ch.Font != null)
                            {
                                if (ch.Font.IsBold.HasValue)
                                    setting.Font.IsBold = ch.Font.IsBold.Value;

                                if (ch.Font.IsItalic.HasValue)
                                    setting.Font.IsItalic = ch.Font.IsItalic.Value;

                                if (!string.IsNullOrEmpty(ch.Font.Color))
                                {
                                    Color color = ColorTranslator.FromHtml(ch.Font.Color);
                                    setting.Font.Color = color;
                                }
                            }

                            fontSettings[i] = setting;
                        }

                        // Apply all character formatting back to the cell.
                        cell.SetCharacters(fontSettings);
                    }
                }
                else
                {
                    Console.WriteLine($"JSON file '{jsonPath}' not found. Skipping rich‑text reapplication.");
                }

                // Save the workbook.
                string outputPath = "ReappliedRichText.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}