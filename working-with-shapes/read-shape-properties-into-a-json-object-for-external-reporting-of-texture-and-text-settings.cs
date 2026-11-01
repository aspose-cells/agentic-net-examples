// Title: Generate a JSON report of Excel shape dimensions, texture fill status, and text formatting using Aspose.Cells for .NET
// Common Searches: C# extract shape geometry and fill type from an Excel workbook with Aspose.Cells | how to serialize Aspose.Cells shape properties to JSON in .NET | retrieve text content and font details of Excel shapes using Aspose.Cells API | export texture fill information of shapes in an .xlsx file via Aspose.Cells | list all shapes on each worksheet and output their attributes to a JSON file in C#
// Tags: Aspose.Cells shape geometry extraction .NET | Excel shape texture fill detection using Aspose.Cells | Aspose.Cells retrieve shape text and font attributes | C# serialize shape collection to JSON file | iterate worksheet shapes with Aspose.Cells API

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet's ShapeCollection, captures each shape's position, size, texture‑fill status, and any embedded text with its font properties, aggregates this data into a list of objects, and writes the list as a formatted JSON file for external reporting.
class ShapeReportGenerator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "ShapeReport.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to hold shape information for all worksheets
            var shapeReport = new List<object>();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                // Process each shape on the worksheet
                foreach (Shape shape in shapes)
                {
                    try
                    {
                        // Basic shape information
                        var shapeInfo = new Dictionary<string, object>
                        {
                            { "Worksheet", sheet.Name },
                            { "Name", shape.Name },
                            { "Type", shape.Type.ToString() },
                            { "Top", shape.Top },
                            { "Left", shape.Left },
                            { "Width", shape.Width },
                            { "Height", shape.Height }
                        };

                        // Fill (texture) information
                        FillFormat fill = shape.Fill;
                        if (fill != null && fill.FillType == FillType.Texture)
                        {
                            // Texture image extraction is optional; Aspose.Cells may not expose the image directly
                            shapeInfo.Add("TextureFill", new Dictionary<string, object>
                            {
                                { "Enabled", true },
                                { "ImageBase64", null } // Placeholder when image data is unavailable
                            });
                        }
                        else
                        {
                            shapeInfo.Add("TextureFill", new Dictionary<string, object>
                            {
                                { "Enabled", false }
                            });
                        }

                        // Text and text formatting (if the shape contains text)
                        if (!string.IsNullOrEmpty(shape.Text))
                        {
                            Font font = shape.Font;
                            shapeInfo.Add("Text", shape.Text);
                            shapeInfo.Add("TextFormatting", new Dictionary<string, object>
                            {
                                { "FontName", font?.Name },
                                { "FontSize", font?.Size },
                                { "Bold", font?.IsBold ?? false },
                                { "Italic", font?.IsItalic ?? false },
                                { "Underline", font?.Underline.ToString() },
                                { "Color", font != null ? font.Color.ToArgb() : (int?)null }
                            });
                        }

                        // Add the collected information to the report list
                        shapeReport.Add(shapeInfo);
                    }
                    catch (Exception shapeEx)
                    {
                        Console.WriteLine($"Warning: Failed to process shape \"{shape.Name}\" on sheet \"{sheet.Name}\": {shapeEx.Message}");
                    }
                }
            }

            // Serialize the report to JSON (using System.Text.Json)
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string jsonReport = JsonSerializer.Serialize(shapeReport, jsonOptions);

            // Write JSON to file
            File.WriteAllText(outputPath, jsonReport);

            Console.WriteLine("Shape report generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
