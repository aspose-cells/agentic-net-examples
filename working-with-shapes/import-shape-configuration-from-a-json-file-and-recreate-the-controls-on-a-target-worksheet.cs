// Title: C# example: Import shape definitions from a JSON file and add them to a specific worksheet using Aspose.Cells
// Common Searches: c# read json configuration and create shapes in aspose.cells worksheet | how to add rectangle and oval shapes to a specific sheet using aspose.cells | example of using MsoDrawingType to insert shapes from json in aspose.cells | load shape definitions from a json file and apply rotation in aspose.cells c# | aspose.cells import shape settings from json and save the workbook
// Tags: Aspose.Cells add shape from JSON | C# Aspose.Cells shape creation API | MsoDrawingType enumeration Aspose.Cells | JSON-driven shape insertion Aspose.Cells | Aspose.Cells worksheet shape positioning | Aspose.Cells shape rotation configuration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace AsposeCellsShapeCreator
{
    // The sample loads an existing workbook, parses a JSON array containing shape definitions (type, position, size, optional text and rotation), creates each shape on the specified worksheet using the Aspose.Cells Shapes.AddShape API, and saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define paths (replace with actual paths or pass via args)
            string workbookPath = @"{WorkbookPath}";
            string jsonConfigPath = @"{JsonConfigPath}";
            string outputWorkbookPath = @"{OutputWorkbookPath}";
            string targetSheetName = @"{TargetSheetName}";

            try
            {
                // Verify workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(workbookPath);

                // Verify JSON config file exists
                if (!File.Exists(jsonConfigPath))
                {
                    Console.WriteLine($"JSON config file not found: {jsonConfigPath}");
                    return;
                }

                // Load and parse JSON configuration
                string jsonContent = File.ReadAllText(jsonConfigPath);
                JsonArray shapeConfigs = JsonNode.Parse(jsonContent)?.AsArray();
                if (shapeConfigs == null)
                {
                    Console.WriteLine("Invalid JSON configuration.");
                    return;
                }

                // Get target worksheet
                Worksheet targetSheet = workbook.Worksheets[targetSheetName];
                if (targetSheet == null)
                {
                    Console.WriteLine($"Worksheet \"{targetSheetName}\" not found.");
                    return;
                }

                // Iterate through each shape definition
                foreach (JsonObject shapeDef in shapeConfigs)
                {
                    try
                    {
                        // Shape type (e.g., "Rectangle", "Oval", "Line", etc.)
                        string typeStr = shapeDef["Type"]?.GetValue<string>();
                        if (string.IsNullOrEmpty(typeStr))
                            continue;

                        if (!Enum.TryParse(typeof(MsoDrawingType), typeStr, true, out object shapeTypeObj))
                        {
                            Console.WriteLine($"Invalid shape type: {typeStr}");
                            continue;
                        }
                        MsoDrawingType shapeType = (MsoDrawingType)shapeTypeObj;

                        // Position and size
                        int upperRow = shapeDef["UpperRow"]?.GetValue<int>() ?? 0;
                        int upperColumn = shapeDef["UpperColumn"]?.GetValue<int>() ?? 0;
                        int upperRowOffset = shapeDef["UpperRowOffset"]?.GetValue<int>() ?? 0;
                        int upperColumnOffset = shapeDef["UpperColumnOffset"]?.GetValue<int>() ?? 0;
                        int height = shapeDef["Height"]?.GetValue<int>() ?? 0;
                        int width = shapeDef["Width"]?.GetValue<int>() ?? 0;

                        // Create shape
                        Shape shape = targetSheet.Shapes.AddShape(
                            shapeType,
                            upperRow,
                            upperRowOffset,
                            upperColumn,
                            upperColumnOffset,
                            height,
                            width);

                        // Optional text
                        if (shapeDef.ContainsKey("Text"))
                        {
                            shape.Text = shapeDef["Text"]?.GetValue<string>();
                        }

                        // Optional rotation angle
                        if (shapeDef.ContainsKey("RotationAngle"))
                        {
                            shape.RotationAngle = shapeDef["RotationAngle"]?.GetValue<double>() ?? 0.0;
                        }

                        // Optional visibility (true = visible, false = hidden)
                        // Note: Aspose.Cells Shape does not expose an IsVisible property; visibility can be controlled via the Visible property of the shape's underlying object if needed.
                        // This block is kept for future compatibility but currently does nothing.
                        if (shapeDef.ContainsKey("Visible"))
                        {
                            // Placeholder for visibility handling if supported in future versions.
                        }
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Error processing a shape definition: {exShape.Message}");
                    }
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputWorkbookPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputWorkbookPath);
                Console.WriteLine($"Workbook saved to: {outputWorkbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
