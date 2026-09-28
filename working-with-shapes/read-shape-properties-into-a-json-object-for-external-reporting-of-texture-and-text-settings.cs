// Title: Extract Excel shape text, font, and fill properties to JSON with Aspose.Cells in C#
// AI Prompts: Generate C# code that opens an .xlsx file using Aspose.Cells, iterates every worksheet and its ShapeCollection, and writes each shape's name, type, text, font attributes (name, size, bold, color) and fill type into a JSON file. | Update the shape export example to also capture each shape's Top, Left, Width, and Height properties and include them in the JSON output using Aspose.Cells.
// Common Searches: how to read shape text and font details from an Excel workbook using Aspose.Cells C# | Aspose.Cells C# export shape fill type and color to JSON | C# iterate worksheet shapes and serialize their properties with System.Text.Json | extract shape metadata name type position from Excel file using Aspose.Cells | save Excel shape information as JSON file in .NET
// Tags: Aspose.Cells extract shape properties to JSON | C# serialize Excel shape metadata | Aspose.Cells shape text and font extraction | export shape fill type Aspose.Cells | read shape position and size Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePropertiesExport
{
    // Class representing the properties we want to export for each shape
    // The example loads an Excel workbook with Aspose.Cells, loops through each worksheet and its ShapeCollection, gathers shape name, type, text, font details (name, size, bold, color) and fill type into a ShapeInfo object, then serializes the list to an indented JSON file named 'shape_properties.json'.
    public class ShapeInfo
    {
        public string? WorksheetName { get; set; }
        public string? ShapeName { get; set; }
        public string? ShapeType { get; set; }

        // Text related properties (if applicable)
        public string? Text { get; set; }
        public string? FontName { get; set; }
        public double? FontSize { get; set; }
        public bool? FontBold { get; set; }
        public string? FontColor { get; set; }

        // Fill/texture related properties (if applicable)
        public string? FillType { get; set; }
        public string? FillForeColor { get; set; }
        public string? FillBackColor { get; set; }
        public string? FillTextureImagePath { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the Excel file to be processed
                string inputPath = "input.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // List to hold shape information from all worksheets
                List<ShapeInfo> shapesInfo = new List<ShapeInfo>();

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the collection of shapes on the current worksheet
                    ShapeCollection shapes = sheet.Shapes;

                    // Iterate through each shape
                    foreach (Shape shape in shapes)
                    {
                        ShapeInfo info = new ShapeInfo
                        {
                            WorksheetName = sheet.Name,
                            ShapeName = shape.Name,
                            ShapeType = shape.Type.ToString()
                        };

                        // ----- Text settings (if the shape supports text) -----
                        if (!string.IsNullOrEmpty(shape.Text))
                        {
                            info.Text = shape.Text;

                            // Font information is available via the Font property
                            Font? font = shape.Font;
                            if (font != null)
                            {
                                info.FontName = font.Name;
                                info.FontSize = font.Size;
                                info.FontBold = font.IsBold;
                                // Convert the font color to a hex string for readability
                                info.FontColor = $"#{font.Color.R:X2}{font.Color.G:X2}{font.Color.B:X2}";
                            }
                        }

                        // ----- Fill settings (if the shape supports fill) -----
                        FillFormat? fill = shape.Fill;
                        if (fill != null)
                        {
                            info.FillType = fill.FillType.ToString();

                            // Note: ForeColor, BackColor, and TextureImage properties are not available
                            // in the current Aspose.Cells version used. If needed, they can be accessed
                            // via alternative APIs in future updates.
                        }

                        shapesInfo.Add(info);
                    }
                }

                // Serialize the collected shape information to JSON using System.Text.Json
                var jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                };
                string jsonOutput = JsonSerializer.Serialize(shapesInfo, jsonOptions);

                // Write JSON to a file
                string outputPath = "shape_properties.json";
                File.WriteAllText(outputPath, jsonOutput);

                Console.WriteLine($"Shape properties have been exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
