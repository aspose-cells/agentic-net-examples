using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePropertiesExport
{
    // DTO to hold shape information
    public class ShapeInfo
    {
        public string? Name { get; set; }
        public string? ShapeType { get; set; }
        public string? Text { get; set; }
        public string? FontName { get; set; }
        public double FontSize { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public string? FillType { get; set; }
        public string? TextureBase64 { get; set; }
    }

    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "shapes.json";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.Error.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load workbook
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.Error.WriteLine($"Error loading workbook: {loadEx.Message}");
                    return;
                }

                // Work with the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Collect shape information
                List<ShapeInfo> shapesInfo = new List<ShapeInfo>();

                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Text and font properties
                        string? text = shape.Text;
                        Font? font = shape.Font;

                        // Fill properties
                        FillFormat fill = shape.Fill;
                        string? fillType = fill.FillType.ToString();

                        // Texture handling – Aspose.Cells may not expose the texture image directly in all versions.
                        // If unavailable, leave TextureBase64 as null.
                        string? textureBase64 = null;

                        // Populate DTO
                        ShapeInfo info = new ShapeInfo
                        {
                            Name = shape.Name,
                            ShapeType = shape.Type.ToString(),
                            Text = text,
                            FontName = font?.Name,
                            FontSize = font?.Size ?? 0,
                            IsBold = font?.IsBold ?? false,
                            IsItalic = font?.IsItalic ?? false,
                            FillType = fillType,
                            TextureBase64 = textureBase64
                        };

                        shapesInfo.Add(info);
                    }
                    catch (Exception shapeEx)
                    {
                        Console.Error.WriteLine($"Warning: Unable to process shape \"{shape.Name}\": {shapeEx.Message}");
                    }
                }

                // Serialize to JSON (indented)
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(shapesInfo, options);

                // Write JSON to output file
                try
                {
                    File.WriteAllText(outputPath, json);
                    Console.WriteLine($"Shape properties have been exported to \"{outputPath}\"");
                }
                catch (Exception writeEx)
                {
                    Console.Error.WriteLine($"Error writing output file: {writeEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}