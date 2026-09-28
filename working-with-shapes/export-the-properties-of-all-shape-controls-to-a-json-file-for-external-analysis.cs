// Title: Export all shape control properties from an Excel workbook to a JSON file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a .xlsx file with Aspose.Cells, iterates through each worksheet and its Shapes collection, captures each shape's name, type, top, left, width, and height, and writes the data to an indented JSON file using System.Text.Json. | Create a reusable method that returns a List of custom ShapeInfo objects for all shapes in a workbook and serializes the list to a JSON string, including error handling for missing input files.
// Common Searches: how to extract shape dimensions from an Excel file using Aspose.Cells C# | C# Aspose.Cells list all shapes and export their properties to JSON | save Excel shape metadata to JSON with System.Text.Json in .NET | iterate worksheets and shapes in Aspose.Cells to generate a JSON report | export shape type and position from workbook using Aspose.Cells API
// Tags: export shape metadata to JSON with Aspose.Cells | iterate workbook shapes using Aspose.Cells API | serialize Excel shape properties using System.Text.Json | collect shape dimensions from .xlsx in .NET | Aspose.Cells shape information extraction

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace AsposeCellsExample
{
    // Class to hold shape properties
    // The example loads an .xlsx workbook with Aspose.Cells, walks through every worksheet and each shape on the sheet, captures the shape's worksheet name, name, type, top, left, width and height into a ShapeInfo object, aggregates them into a list, serializes the list to formatted JSON using System.Text.Json, and writes the result to shape_properties.json while handling missing files and exceptions.
    class ShapeInfo
    {
        public string? Worksheet { get; set; }
        public string? Name { get; set; }
        public string? ShapeType { get; set; }
        public double Top { get; set; }
        public double Left { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "shape_properties.json";

                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                var shapeData = new List<ShapeInfo>();

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each shape in the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Collect shape information
                        shapeData.Add(new ShapeInfo
                        {
                            Worksheet = sheet.Name,
                            Name = shape.Name,
                            ShapeType = shape.Type.ToString(),
                            Top = shape.Top,
                            Left = shape.Left,
                            Width = shape.Width,
                            Height = shape.Height
                        });
                    }
                }

                // Convert the collected data to JSON
                string json = JsonSerializer.Serialize(shapeData, new JsonSerializerOptions { WriteIndented = true });

                // Save the JSON to a file
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"Shape properties saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
