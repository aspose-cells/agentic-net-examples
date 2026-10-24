// Title: Create a rectangle shape in Aspose.Cells, set its absolute position from JSON pixel coordinates, and verify placement using C#
// Common Searches: Aspose.Cells C# set shape top and left using pixel values from a JSON file | how to position a free‑floating shape in an Excel workbook based on configuration data | convert pixel coordinates to points for shape placement in Aspose.Cells .NET | validate that a rectangle shape is positioned correctly after applying absolute coordinates | read shape X Y coordinates from config.json and apply to a shape in Aspose.Cells
// Tags: set shape absolute position Aspose.Cells | pixel to point conversion Aspose.Cells | load shape coordinates from JSON C# | free floating shape placement Aspose.Cells | verify shape top left values Aspose.Cells

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example reads X and Y pixel coordinates from a JSON configuration file, creates a new workbook, adds a rectangle shape at cell A1, converts the pixel values to points (0.75 points per pixel), sets the shape's placement to FreeFloating, applies the calculated Top and Left values, validates that the shape's position matches the expected points, and saves the workbook as Output.xlsx.
class Config
{
    public int ShapePixelX { get; set; }
    public int ShapePixelY { get; set; }
}

class Program
{
    static void Main()
    {
        try
        {
            // Load pixel configuration from a JSON file.
            const string configPath = "config.json";
            if (!File.Exists(configPath))
            {
                Console.WriteLine($"Configuration file '{configPath}' not found.");
                return;
            }

            string json = File.ReadAllText(configPath);
            Config? cfg = JsonSerializer.Deserialize<Config>(json);
            if (cfg == null)
            {
                Console.WriteLine("Failed to deserialize configuration.");
                return;
            }

            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape at cell A1 with temporary zero offsets.
            const int initialRow = 0;
            const int initialColumn = 0;
            const int shapeWidth = 100;   // points
            const int shapeHeight = 50;   // points
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                initialRow,
                initialColumn,
                0,
                0,
                shapeWidth,
                shapeHeight);

            // Convert pixel values to points (1 pixel = 0.75 points at 96 DPI).
            const double pixelsToPoints = 72.0 / 96.0; // 0.75
            int topPoints = (int)Math.Round(cfg.ShapePixelY * pixelsToPoints);
            int leftPoints = (int)Math.Round(cfg.ShapePixelX * pixelsToPoints);

            // Set the shape to free‑floating placement and apply absolute position.
            shape.Placement = PlacementType.FreeFloating;
            shape.Top = topPoints;
            shape.Left = leftPoints;

            // Validate that the shape's position matches the expected values.
            bool topValid = Math.Abs(shape.Top - topPoints) < 0.001;
            bool leftValid = Math.Abs(shape.Left - leftPoints) < 0.001;
            Console.WriteLine(topValid && leftValid
                ? "Shape positioned correctly."
                : "Shape positioning mismatch.");

            // Save the workbook.
            const string outputPath = "Output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
