// Title: Export Excel shape properties (position, size, fill, line style) to JSON using Aspose.Cells for .NET
// Common Searches: how to extract shape coordinates and dimensions from an Excel file with Aspose.Cells C# | save Aspose.Cells worksheet shape details as JSON in .NET | C# get shape fill color and border style from Excel using Aspose.Cells | export all shapes metadata from multiple worksheets to a JSON file with Aspose.Cells | retrieve shape rotation and Z-order using Aspose.Cells API in C#
// Tags: Aspose.Cells export shape metadata to JSON | C# extract Excel shape properties | Aspose.Cells shape position and size | Aspose.Cells shape fill color serialization | Aspose.Cells line format to JSON

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeMetadataExport
{
    // POCO to hold shape metadata that will be serialized to JSON
    // The example loads an Excel workbook with Aspose.Cells, iterates through every worksheet and its ShapeCollection, and captures detailed metadata for each shape—including worksheet name, shape name, alternative text, title, auto‑shape type, rotation angle, Z‑order, pixel‑based left/top coordinates, width, height, solid fill color (ARGB), line weight, dash style, and line color (when available). The collected data is stored in a POCO list, serialized to an indented JSON string using System.Text.Json, and written to a user‑specified output file. The code includes command‑line path overrides, file existence checks, and graceful handling of shape‑specific exceptions.
    public class ShapeInfo
    {
        public string? Worksheet { get; set; }
        public string? Name { get; set; }
        public string? AlternativeText { get; set; }
        public string? Title { get; set; }
        public AutoShapeType? AutoShapeType { get; set; }
        public double RotationAngle { get; set; }
        public int ZOrderPosition { get; set; }

        // Position (in pixels)
        public double Left { get; set; }
        public double Top { get; set; }

        // Size (in pixels)
        public double Width { get; set; }
        public double Height { get; set; }

        // Fill information (only solid fill is captured)
        public string? FillColor { get; set; }

        // Line (border) information
        public double LineWeight { get; set; }
        public string? LineDashStyle { get; set; }
        public string? LineColor { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (can be overridden via command‑line)
            string inputPath = "input.xlsx";
            string outputPath = "shapes.json";

            if (args.Length >= 1) inputPath = args[0];
            if (args.Length >= 2) outputPath = args[1];

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.Error.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Collect shape metadata from all worksheets
                List<ShapeInfo> allShapes = new List<ShapeInfo>();

                for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
                {
                    Worksheet ws = workbook.Worksheets[wsIndex];
                    ShapeCollection shapes = ws.Shapes;

                    for (int shapeIndex = 0; shapeIndex < shapes.Count; shapeIndex++)
                    {
                        Shape shape = shapes[shapeIndex];
                        try
                        {
                            var info = new ShapeInfo
                            {
                                Worksheet = ws.Name,
                                Name = shape.Name,
                                AlternativeText = shape.AlternativeText,
                                Title = shape.Title,
                                AutoShapeType = shape.AutoShapeType,
                                RotationAngle = shape.RotationAngle,
                                ZOrderPosition = shape.ZOrderPosition,

                                // Position
                                Left = shape.Left,
                                Top = shape.Top,

                                // Size
                                Width = shape.Width,
                                Height = shape.Height,

                                // Fill (only solid fill is handled)
                                FillColor = shape.Fill.FillType == FillType.Solid
                                    ? shape.Fill.SolidFill.Color.ToArgb().ToString("X8")
                                    : null,

                                // Line (border) details
                                LineWeight = shape.Line.Weight,
                                LineDashStyle = shape.Line.DashStyle.ToString(),
                                // Line color may not be available in some versions; guard against it
                                LineColor = GetLineColor(shape.Line)
                            };

                            allShapes.Add(info);
                        }
                        catch (Exception shapeEx)
                        {
                            Console.Error.WriteLine($"Warning: Failed to process shape \"{shape.Name}\" on sheet \"{ws.Name}\": {shapeEx.Message}");
                        }
                    }
                }

                // Serialize the list to JSON with indentation for readability
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(allShapes, jsonOptions);

                // Write JSON to the output file
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"Successfully exported {allShapes.Count} shape(s) to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }

        // Helper to safely obtain line color as ARGB string; returns null if not supported
        private static string? GetLineColor(LineFormat line)
        {
            try
            {
                // Some versions expose Color, others may not; use reflection as a fallback
                var colorProp = typeof(LineFormat).GetProperty("Color");
                if (colorProp != null)
                {
                    var colorValue = colorProp.GetValue(line);
                    if (colorValue is Color sysColor)
                    {
                        return sysColor.ToArgb().ToString("X8");
                    }
                }
            }
            catch
            {
                // Ignored – return null if any issue occurs
            }
            return null;
        }
    }
}
