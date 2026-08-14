// Title: Reconstruct SmartArt from GroupShape JSON using Aspose.Cells (C#)
// Description: Loads a JSON file that describes rectangle and oval shapes, creates matching Aspose.Cells shapes, converts flagged items to SmartArt, groups them to restore the original hierarchy, and saves the workbook as an Excel file.
// Keywords: Aspose.Cells | C# | SmartArt reconstruction | GroupShape JSON | shape deserialization | Excel shape hierarchy | rectangle shape | oval shape | GetResultOfSmartArt | group shapes | workbook save
// Common Searches: Aspose.Cells deserialize GroupShape JSON | C# rebuild SmartArt from JSON | create and group shapes in Excel with Aspose.Cells | convert JSON shape data to SmartArt hierarchy | load shape definitions from JSON into workbook
// Developer Intent: Restore a saved SmartArt layout in an Excel workbook by deserializing a GroupShape JSON file and programmatically recreating and grouping the shapes with Aspose.Cells.
// Use Cases: Read a JSON file containing shape type, position, size, text, and SmartArt flag, then generate the corresponding Rectangle or Oval objects on a worksheet. | Detect shapes marked as SmartArt, retrieve their SmartArt representation via GetResultOfSmartArt, and replace the placeholder shape with the actual SmartArt group. | Group all generated shapes into a single GroupShape to re‑establish the original hierarchy, with an optional ungroup step for later editing. | Handle missing or empty JSON files gracefully, ensuring a valid workbook is still produced.
// AI Prompts: Generate C# code that reads a GroupShapeModel JSON file and uses Aspose.Cells to recreate each shape and the SmartArt hierarchy in an Excel workbook. | Provide a mapping function that converts ShapeModel properties (type, row, column, width, height, text) to the appropriate Aspose.Cells shape‑creation calls, including fallback for unknown types. | Explain how to group an array of Shape objects into a GroupShape, optionally ungroup it later, and save the resulting workbook.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtReconstruction
{
    // Model representing a single shape saved in JSON
    // Loads a JSON file that describes rectangle and oval shapes, creates matching Aspose.Cells shapes, converts flagged items to SmartArt, groups them to restore the original hierarchy, and saves the workbook as an Excel file.
    public class ShapeModel
    {
        public string ShapeType { get; set; }          // e.g., "Rectangle", "Oval"
        public int UpperLeftRow { get; set; }
        public int UpperLeftColumn { get; set; }
        public int Height { get; set; }                // in pixels
        public int Width { get; set; }                 // in pixels
        public string Text { get; set; }               // optional text
        public bool IsSmartArt { get; set; }           // indicates original SmartArt
    }

    // Model representing the saved GroupShape hierarchy
    public class GroupShapeModel
    {
        public List<ShapeModel> Shapes { get; set; } = new List<ShapeModel>();
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // ---------- Create / Load Workbook ----------
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ---------- Read and Deserialize JSON ----------
                string jsonPath = "savedGroupShape.json";
                GroupShapeModel groupModel;

                if (File.Exists(jsonPath))
                {
                    string jsonContent = File.ReadAllText(jsonPath);
                    groupModel = JsonSerializer.Deserialize<GroupShapeModel>(jsonContent);
                    if (groupModel == null)
                        groupModel = new GroupShapeModel();
                }
                else
                {
                    // If the JSON file is missing, continue with an empty model
                    Console.WriteLine($"Warning: JSON file '{jsonPath}' not found. Proceeding with empty shape collection.");
                    groupModel = new GroupShapeModel();
                }

                // List to hold the created Shape objects before grouping
                List<Shape> createdShapes = new List<Shape>();

                // ---------- Recreate Individual Shapes ----------
                foreach (var shapeInfo in groupModel.Shapes)
                {
                    Shape shape = null;

                    // Create shape based on the stored type (Rectangle or Oval)
                    if (shapeInfo.ShapeType.Equals("Rectangle", StringComparison.OrdinalIgnoreCase))
                    {
                        shape = worksheet.Shapes.AddRectangle(
                            shapeInfo.UpperLeftRow,
                            shapeInfo.UpperLeftColumn,
                            0, // offset X in pixels
                            0, // offset Y in pixels
                            shapeInfo.Width,
                            shapeInfo.Height);
                    }
                    else if (shapeInfo.ShapeType.Equals("Oval", StringComparison.OrdinalIgnoreCase))
                    {
                        shape = worksheet.Shapes.AddOval(
                            shapeInfo.UpperLeftRow,
                            shapeInfo.UpperLeftColumn,
                            0,
                            0,
                            shapeInfo.Width,
                            shapeInfo.Height);
                    }
                    else
                    {
                        // Fallback to a generic rectangle if type is unknown
                        shape = worksheet.Shapes.AddRectangle(
                            shapeInfo.UpperLeftRow,
                            shapeInfo.UpperLeftColumn,
                            0,
                            0,
                            shapeInfo.Width,
                            shapeInfo.Height);
                    }

                    // Apply optional text
                    if (!string.IsNullOrEmpty(shapeInfo.Text))
                        shape.Text = shapeInfo.Text;

                    // Attempt to convert original SmartArt shapes
                    if (shapeInfo.IsSmartArt && shape.IsSmartArt)
                    {
                        GroupShape smartArtGroup = shape.GetResultOfSmartArt();
                        if (smartArtGroup != null)
                            shape = smartArtGroup;
                    }

                    createdShapes.Add(shape);
                }

                // ---------- Group Shapes to Recreate Hierarchy ----------
                if (createdShapes.Count > 0)
                {
                    Shape[] shapesArray = createdShapes.ToArray();
                    GroupShape reconstructedGroup = worksheet.Shapes.Group(shapesArray);
                    // Optional: Ungroup if you need individual access later
                    // reconstructedGroup.Ungroup();
                }

                // ---------- Save Workbook ----------
                string outputPath = "ReconstructedSmartArt.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
