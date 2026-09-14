// Title: Serialize the RichTextPortion collection of an Aspose.Cells cell to indented JSON in C#
// AI Prompts: Generate C# code that extracts all FontSetting objects from a worksheet cell and maps their properties (start index, length, bold, italic, underline, font name, size, color) into a plain DTO list. | Create a System.Text.Json JsonConverter that writes System.Drawing.Color values as #AARRGGBB hexadecimal strings for inclusion in the serialized output. | Combine the DTO list and the custom color converter to produce a formatted (indented) JSON string representing the cell's rich‑text portions.
// Common Searches: Aspose.Cells C# export cell rich text formatting to JSON | How to get FontSetting objects from a cell and serialize them with System.Text.Json | Custom JsonConverter for System.Drawing.Color hex string in Aspose.Cells example | Serialize rich text portions of a worksheet cell to JSON using Aspose.Cells .NET
// Tags: Aspose.Cells serialize rich text portions to JSON | C# System.Text.Json custom color converter | FontSetting to DTO mapping Aspose.Cells | rich text formatting JSON export .NET | cell characters JSON serialization Aspose

using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsRichTextSerialization
{
    // Custom converter to serialize System.Drawing.Color as a hex string
    // Shows how to retrieve FontSetting objects from a cell, map them to a lightweight DTO, and serialize the collection to indented JSON using System.Text.Json with a custom converter that outputs System.Drawing.Color as a #AARRGGBB hex string.
    public class ColorJsonConverter : JsonConverter<Color>
    {
        public override Color Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            // Deserialization not required for this example
            throw new NotImplementedException();
        }

        public override void Write(Utf8JsonWriter writer, Color value, JsonSerializerOptions options)
        {
            // Serialize as #AARRGGBB
            string hex = $"#{value.A:X2}{value.R:X2}{value.G:X2}{value.B:X2}";
            writer.WriteStringValue(hex);
        }
    }

    // Helper class to hold serializable information of a rich text portion
    public class RichTextPortionDto
    {
        public int StartIndex { get; set; }
        public int Length { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public string FontName { get; set; }
        public double FontSize { get; set; }

        [JsonConverter(typeof(ColorJsonConverter))]
        public Color FontColor { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a cell with rich text formatting
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("Hello World");

            // Format "Hello" as bold red
            FontSetting part1 = cell.Characters(0, 5);
            part1.Font.IsBold = true;
            part1.Font.Color = Color.Red;

            // Format "World" as italic blue
            FontSetting part2 = cell.Characters(6, 5);
            part2.Font.IsItalic = true;
            part2.Font.Color = Color.Blue;

            // Retrieve all rich text portions (FontSetting objects)
            FontSetting[] portions = cell.GetCharacters();

            // Convert FontSetting objects to DTOs suitable for JSON serialization
            List<RichTextPortionDto> dtoList = new List<RichTextPortionDto>();
            foreach (FontSetting fs in portions)
            {
                dtoList.Add(new RichTextPortionDto
                {
                    StartIndex = fs.StartIndex,
                    Length = fs.Length,
                    IsBold = fs.Font.IsBold,
                    IsItalic = fs.Font.IsItalic,
                    IsUnderline = fs.Font.Underline != FontUnderlineType.None,
                    FontName = fs.Font.Name,
                    FontSize = fs.Font.Size,
                    FontColor = fs.Font.Color
                });
            }

            // Serialize the collection to JSON with indentation
            JsonSerializerOptions jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true,
                Converters = { new ColorJsonConverter() }
            };
            string json = JsonSerializer.Serialize(dtoList, jsonOptions);

            // Output the JSON string
            Console.WriteLine("Rich Text Portions JSON:");
            Console.WriteLine(json);
        }
    }
}
