using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsRichTextSerialization
{
    // DTO that represents a rich‑text portion of a cell
    public class RichTextPortionDto
    {
        public int StartIndex { get; set; }
        public int Length { get; set; }

        // Font related properties we want to preserve
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public string Color { get; set; }   // stored as ARGB hex string
        public double FontSize { get; set; }
        public string FontName { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a workbook and add rich‑text to a cell
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Hello World!");                         // base text

            // Apply formatting to two portions of the text
            // "Hello" (0‑4) – bold red
            FontSetting part1 = cell.Characters(0, 5);
            part1.Font.IsBold = true;
            part1.Font.Color = Color.Red;
            part1.Font.Size = 12;
            part1.Font.Name = "Arial";

            // "World!" (6‑11) – italic blue
            FontSetting part2 = cell.Characters(6, 6);
            part2.Font.IsItalic = true;
            part2.Font.Color = Color.Blue;
            part2.Font.Size = 12;
            part2.Font.Name = "Calibri";

            // -------------------------------------------------
            // 2. Retrieve the rich‑text portions (FontSetting[])
            // -------------------------------------------------
            FontSetting[] portions = cell.GetCharacters();

            // -------------------------------------------------
            // 3. Convert each FontSetting to a serializable DTO
            // -------------------------------------------------
            List<RichTextPortionDto> dtoList = new List<RichTextPortionDto>();
            foreach (FontSetting fs in portions)
            {
                RichTextPortionDto dto = new RichTextPortionDto
                {
                    StartIndex = fs.StartIndex,
                    Length = fs.Length,
                    IsBold = fs.Font.IsBold,
                    IsItalic = fs.Font.IsItalic,
                    IsUnderline = fs.Font.Underline != FontUnderlineType.None,
                    Color = ColorTranslator.ToHtml(fs.Font.Color), // store as HTML hex (e.g., "#FF0000")
                    FontSize = fs.Font.Size,
                    FontName = fs.Font.Name
                };
                dtoList.Add(dto);
            }

            // -------------------------------------------------
            // 4. Serialize the DTO list to JSON
            // -------------------------------------------------
            JsonSerializerOptions jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            string json = JsonSerializer.Serialize(dtoList, jsonOptions);

            // -------------------------------------------------
            // 5. Output the JSON (could also be saved to a file)
            // -------------------------------------------------
            Console.WriteLine("Rich‑text portions JSON:");
            Console.WriteLine(json);

            // Optional: save JSON to a file for later reuse
            System.IO.File.WriteAllText("RichTextPortions.json", json);

            // -------------------------------------------------
            // 6. Save the workbook (demonstrates normal lifecycle)
            // -------------------------------------------------
            workbook.Save("RichTextDemo.xlsx");
        }
    }
}