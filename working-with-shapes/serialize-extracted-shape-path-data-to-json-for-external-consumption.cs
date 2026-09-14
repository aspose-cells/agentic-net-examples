// Title: Export Excel shape properties (name, alt text, type) to indented JSON using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, iterates through every worksheet and shape, extracts each shape's Name, AlternativeText, and Type, and saves the collected data as formatted JSON to a target file. | Enhance the previous solution to also capture each shape's Top, Left, Width, and Height values and include these dimensions in the generated JSON output.
// Common Searches: how to extract shape information from Excel using Aspose.Cells C# | Aspose.Cells C# export shape name and type to JSON file | serialize Excel drawing objects to JSON with Aspose.Cells .NET
// Tags: Aspose.Cells shape metadata JSON export | C# collect Excel shape properties | serialize shape position size Aspose.Cells | write formatted JSON from shape list C# | extract shape type enumeration to JSON

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// Loads input.xlsx with Aspose.Cells, iterates all worksheets and shapes, gathers each shape's Name, AlternativeText, and Type, then writes the list as indented JSON to shapes.json.
class ShapeInfo
{
    public string Name { get; set; } = string.Empty;
    public string AlternativeText { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
}

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "shapes.json";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect shape information from all worksheets
            List<ShapeInfo> shapeData = new List<ShapeInfo>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    shapeData.Add(new ShapeInfo
                    {
                        Name = shape.Name ?? string.Empty,
                        AlternativeText = shape.AlternativeText ?? string.Empty,
                        Type = shape.Type.ToString()
                    });
                }
            }

            // Serialize the collected data to JSON
            string json = JsonSerializer.Serialize(shapeData, new JsonSerializerOptions { WriteIndented = true });

            // Write JSON to file
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"Shape data successfully written to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading workbook, accessing shapes, file I/O)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
