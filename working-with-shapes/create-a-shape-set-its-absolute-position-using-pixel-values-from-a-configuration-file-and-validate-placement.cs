// Title: Add a rectangle shape to an Excel workbook with Aspose.Cells, set its absolute pixel position from a JSON configuration file, and verify the placement (C#)
// Common Searches: Aspose.Cells C# set shape left and top using pixel values from JSON | read pixel coordinates from a config file and position a rectangle shape in an Excel workbook with Aspose.Cells | convert pixel offsets to points for shape placement in Aspose.Cells C# | verify shape coordinates after applying absolute position in Aspose.Cells | load JSON configuration and apply absolute shape position in an Aspose.Cells workbook
// Tags: Aspose.Cells rectangle shape absolute positioning | pixel to point conversion Aspose.Cells | C# read JSON configuration for shape coordinates | validate shape placement Aspose.Cells | save workbook with positioned shape Aspose.Cells

using System;
using System.IO;
using System.Text.Json;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePositionExample
{
    // Simple class to hold pixel configuration
    // The example reads X and Y pixel offsets from a JSON file, creates a new workbook, adds a rectangle shape, converts the pixel values to points (using a 0.75 factor for 96 DPI), assigns the calculated points to the shape's Left and Top properties, asserts that the converted values match the original pixel offsets, and saves the workbook as an XLSX file.
    public class PositionConfig
    {
        public int X { get; set; }   // Horizontal pixel offset
        public int Y { get; set; }   // Vertical pixel offset
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Load pixel values from configuration file ----------
                // Expected JSON format: { "X": 100, "Y": 150 }
                string configPath = "config.json";
                if (!File.Exists(configPath))
                    throw new FileNotFoundException($"Configuration file not found: {configPath}");

                string json = File.ReadAllText(configPath);
                PositionConfig config = JsonSerializer.Deserialize<PositionConfig>(json)
                                          ?? throw new InvalidOperationException("Failed to parse configuration.");

                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();               // create workbook
                Worksheet sheet = workbook.Worksheets[0];        // use first worksheet

                // ---------- Add a rectangle shape ----------
                // Parameters: shape type, upper left row, upper left column,
                // upper left row offset, upper left column offset, height, width (all in points)
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    0,               // upper left row
                    0,               // upper left column
                    0,               // row offset (pixels)
                    0,               // column offset (pixels)
                    50,              // height (points)
                    100);            // width (points)

                // ---------- Set absolute position using pixel values ----------
                // Aspose.Cells uses points for Left/Top. Convert pixels to points (1 point = 1/72 inch,
                // assuming 96 DPI, 1 pixel ≈ 0.75 point). Adjust conversion as needed.
                const double pixelsToPoints = 0.75;
                shape.Left = (int)Math.Round(config.X * pixelsToPoints);
                shape.Top = (int)Math.Round(config.Y * pixelsToPoints);

                // ---------- Validate placement ----------
                // Retrieve the actual position (in points) and convert back to pixels for verification.
                int actualX = (int)Math.Round(shape.Left / pixelsToPoints);
                int actualY = (int)Math.Round(shape.Top / pixelsToPoints);

                Debug.Assert(actualX == config.X, $"Horizontal position mismatch: expected {config.X}, got {actualX}");
                Debug.Assert(actualY == config.Y, $"Vertical position mismatch: expected {config.Y}, got {actualY}");

                // ---------- Save the workbook ----------
                string outputPath = "ShapePositionDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);      // save workbook

                Console.WriteLine($"Workbook saved to '{outputPath}'. Shape positioned at ({actualX}px, {actualY}px).");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File error: {ex.Message}");
            }
            catch (JsonException ex)
            {
                Console.Error.WriteLine($"JSON parsing error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
