using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeExport
{
    // DTO to hold shape information that will be serialized to JSON
    public class ShapeInfo
    {
        public string WorksheetName { get; set; }
        public string ShapeName { get; set; }
        // Store shape type as string to avoid enum reference issues
        public string ShapeType { get; set; }
        public bool IsHidden { get; set; }
        // ControlData is binary; represent it as a Base64 string for JSON readability
        public string ControlDataBase64 { get; set; }
    }

    public class ExportShapeControls
    {
        public static void Run()
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create (or load) a workbook. Here we create a new workbook for demo.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook(); // create new workbook
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DemoSheet";

                // Add a sample shape (e.g., a rectangle). In real scenarios shapes may already exist.
                Shape shape = sheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);
                shape.Name = "SampleRectangle";

                // -----------------------------------------------------------------
                // 2. Collect properties of all shapes across all worksheets.
                // -----------------------------------------------------------------
                List<ShapeInfo> shapeInfos = new List<ShapeInfo>();

                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (Shape shp in ws.Shapes)
                    {
                        // Retrieve ControlData (read‑only byte array). It may be null for non‑control shapes.
                        byte[] controlData = shp.ControlData;

                        shapeInfos.Add(new ShapeInfo
                        {
                            WorksheetName = ws.Name,
                            ShapeName = shp.Name,
                            ShapeType = shp.Type.ToString(),
                            IsHidden = shp.IsHidden,
                            ControlDataBase64 = controlData != null ? Convert.ToBase64String(controlData) : null
                        });
                    }
                }

                // -----------------------------------------------------------------
                // 3. Serialize the collected information to JSON.
                // -----------------------------------------------------------------
                JsonSerializerOptions jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                string json = JsonSerializer.Serialize(shapeInfos, jsonOptions);

                // -----------------------------------------------------------------
                // 4. Save the JSON string to a file.
                // -----------------------------------------------------------------
                string outputPath = "shape_controls.json";

                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllText(outputPath, json);
                Console.WriteLine($"Shape control data exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            ExportShapeControls.Run();
        }
    }
}