// Title: Export the rich‑text formatting of a single Excel cell to a JSON file using Aspose.Cells for .NET
// Common Searches: how to extract cell rich text formatting to JSON with Aspose.Cells C# | save Excel cell A1 rich text styles as JSON using Aspose.Cells | serialize Aspose.Cells RichTextPortion objects to JSON in .NET | C# code to write cell rich‑text attributes to a JSON file | convert Excel cell rich text to JSON with Aspose.Cells library
// Tags: Aspose.Cells export cell rich text to JSON | C# serialize RichTextPortion collection | Excel cell formatting JSON serialization Aspose | RichTextPortion DTO creation Aspose.Cells | write cell rich text attributes to file .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example loads an Excel workbook, reads the text and font properties of cell A1, maps them into a RichTextPortionDto list, serializes the list to indented JSON with System.Text.Json, and saves the output to a file named CellRichText.json.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputJson = "CellRichText.json";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Get the target cell
            Cell cell = sheet.Cells["A1"];

            // Retrieve cell text and style (single rich‑text portion)
            string cellText = cell.StringValue; // Use StringValue property instead of GetStringValue
            Style cellStyle = cell.GetStyle();
            Font cellFont = cellStyle.Font;

            // Prepare DTO list (single portion in this simple case)
            List<RichTextPortionDto> portions = new List<RichTextPortionDto>
            {
                new RichTextPortionDto
                {
                    Text = cellText,
                    FontName = cellFont.Name,
                    FontSize = cellFont.Size,
                    IsBold = cellFont.IsBold,
                    IsItalic = cellFont.IsItalic,
                    ColorArgb = cellFont.Color.ToArgb()
                }
            };

            // Serialize DTO list to formatted JSON
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(portions, jsonOptions);

            // Write JSON to file
            File.WriteAllText(outputJson, json);
            Console.WriteLine($"Rich‑text data saved to \"{outputJson}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// DTO representing a rich‑text portion in a JSON‑friendly format
class RichTextPortionDto
{
    public string Text { get; set; } = string.Empty;
    public string FontName { get; set; } = string.Empty;
    public double FontSize { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public int ColorArgb { get; set; }
}
