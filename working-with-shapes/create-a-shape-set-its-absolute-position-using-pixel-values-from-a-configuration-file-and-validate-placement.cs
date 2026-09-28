// Title: Add a rectangle shape to an Excel worksheet with Aspose.Cells, position it using pixel offsets from a JSON configuration, and verify its placement
// AI Prompts: Load X and Y pixel offsets from a JSON file and use Aspose.Cells to insert a rectangle shape at those pixel coordinates on the first worksheet. | Check that the shape's UpperLeftRow and UpperLeftColumn are both zero after insertion and log a success or error message. | Save the workbook as an XLSX file while handling errors from missing configuration, JSON deserialization, shape creation, or file saving.
// Common Searches: Aspose.Cells C# set shape top and left offset in pixels from JSON file | How to validate UpperLeftRow and UpperLeftColumn of a shape after adding it with Aspose.Cells | Read shape coordinates from a configuration file and place shape in Excel using Aspose.Cells | Exception handling for missing or malformed shape configuration when adding a shape in Aspose.Cells
// Tags: Aspose.Cells add rectangle shape pixel positioning | read shape coordinates from JSON C# | validate shape UpperLeftRow UpperLeftColumn Aspose.Cells | save workbook to XLSX after shape insertion | error handling for shape configuration file

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example reads X/Y pixel offsets from a JSON configuration file, creates a new workbook, adds a rectangle shape at row 0 column 0 with the specified pixel offsets, verifies that the shape's UpperLeftRow and UpperLeftColumn are zero, and saves the workbook as an XLSX file, all with comprehensive error handling.
public class ShapeConfig
{
    public int X { get; set; } // Horizontal offset in pixels
    public int Y { get; set; } // Vertical offset in pixels
}

class Program
{
    static void Main()
    {
        try
        {
            // ----- Load configuration -----
            // Expected file format (JSON): { "X": 150, "Y": 80 }
            string configPath = "shapeConfig.json";
            if (!File.Exists(configPath))
            {
                Console.WriteLine($"Configuration file not found: {configPath}");
                return;
            }

            ShapeConfig? config;
            try
            {
                string json = File.ReadAllText(configPath);
                config = JsonSerializer.Deserialize<ShapeConfig>(json);
                if (config == null)
                {
                    Console.WriteLine("Failed to deserialize configuration.");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading configuration: {ex.Message}");
                return;
            }

            // ----- Create a new workbook -----
            Workbook workbook = new Workbook();               // Create a new workbook
            Worksheet sheet = workbook.Worksheets[0];        // Use the first worksheet

            // ----- Add a rectangle shape -----
            // Parameters: drawing type, upper left row, upper left column,
            // top offset (pixels), left offset (pixels), height (pixels), width (pixels)
            const int defaultHeight = 100; // arbitrary height
            const int defaultWidth = 200;  // arbitrary width

            Aspose.Cells.Drawing.Shape shape;
            try
            {
                shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // Shape type
                    0,                        // Upper left row index
                    0,                        // Upper left column index
                    config.Y,                 // Top offset in pixels (vertical)
                    config.X,                 // Left offset in pixels (horizontal)
                    defaultHeight,            // Height in pixels
                    defaultWidth);            // Width in pixels
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding shape: {ex.Message}");
                return;
            }

            // Optional: give the shape a name for later reference
            shape.Name = "ConfiguredRectangle";

            // ----- Validate placement (row/column only) -----
            int row = shape.UpperLeftRow;
            int column = shape.UpperLeftColumn;

            bool isValid = (row == 0) && (column == 0);
            if (!isValid)
            {
                Console.WriteLine("Shape placement validation failed.");
                Console.WriteLine($"Expected - Row:0, Column:0");
                Console.WriteLine($"Actual   - Row:{row}, Column:{column}");
                return;
            }
            else
            {
                Console.WriteLine("Shape placed correctly according to configuration.");
            }

            // ----- Save the workbook -----
            string outputPath = "ShapePositioned.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
