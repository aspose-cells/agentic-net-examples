// Title: Export all Excel shape control properties to a JSON file with Aspose.Cells for .NET (C#)
// Common Searches: how to extract form control properties from an Excel workbook using Aspose.Cells in C# | C# Aspose.Cells export shape controls to a JSON file | list checkbox and dropdown metadata in an Excel file with Aspose.Cells | serialize Aspose.Cells shape objects to JSON for external analysis | retrieve Excel form control attributes via reflection using Aspose.Cells .NET
// Tags: Aspose.Cells shape control metadata extraction | C# serialize Excel shape properties to JSON | Aspose.Cells .NET form control property capture | Excel worksheet shape enumeration with Aspose.Cells | reflection-based shape property collection Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeControlsExport
{
    // DTO to hold shape control information
    // The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet and its shapes, filters for form control shapes, captures all readable properties of each control via reflection into a dictionary, and writes the collected data as an indented JSON file for external analysis.
    public class ShapeControlInfo
    {
        public string Worksheet { get; set; } = string.Empty;
        public string ShapeName { get; set; } = string.Empty;
        public string ShapeType { get; set; } = string.Empty;

        // Generic properties captured via reflection
        public Dictionary<string, object?> Properties { get; set; } = new Dictionary<string, object?>();
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ShapeControlsExport <input_excel_file> [output_json_file]");
                return;
            }

            string inputPath = args[0];
            string outputPath = args.Length > 1 ? args[1] : "shape_controls.json";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                List<ShapeControlInfo> controls = new List<ShapeControlInfo>();

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Determine if the shape is a control based on its runtime type
                        if (!IsControlShape(shape))
                            continue;

                        ShapeControlInfo info = new ShapeControlInfo
                        {
                            Worksheet = sheet.Name,
                            ShapeName = shape.Name,
                            ShapeType = shape.Type.ToString()
                        };

                        // Capture generic properties via reflection (excluding indexers and methods)
                        var props = shape.GetType().GetProperties();
                        foreach (var prop in props)
                        {
                            // Skip indexers or non‑readable properties
                            if (prop.GetIndexParameters().Length > 0 || !prop.CanRead)
                                continue;

                            try
                            {
                                object? value = prop.GetValue(shape);
                                // Store the raw value; if it's not directly serializable, its ToString() will be used by JSON serializer
                                info.Properties[prop.Name] = value;
                            }
                            catch (Exception ex)
                            {
                                // Store the exception message for problematic properties
                                info.Properties[prop.Name] = $"<Error: {ex.Message}>";
                            }
                        }

                        controls.Add(info);
                    }
                }

                // Serialize the list to JSON with indentation for readability
                var jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
                };

                string json = JsonSerializer.Serialize(controls, jsonOptions);
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"Successfully exported {controls.Count} shape control(s) to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        // Helper method to decide whether a shape is a control
        private static bool IsControlShape(Shape shape)
        {
            // Runtime type checks for specific control classes
            return shape is CheckBox ||
                   shape is ComboBox ||
                   shape is ListBox ||
                   shape is RadioButton ||
                   shape is Button ||
                   shape is GroupBox ||
                   shape is Label ||
                   shape is TextBox;
        }
    }
}
