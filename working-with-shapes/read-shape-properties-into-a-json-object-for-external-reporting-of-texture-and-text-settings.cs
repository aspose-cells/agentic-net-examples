// Title: Export Excel Shape Text, Font, Fill & Texture to JSON with Aspose.Cells (C#)
// Description: A C# utility that loads an Excel workbook using Aspose.Cells, walks through every worksheet, extracts each shape's properties—including name, type, text, alignment, rich‑text flag, alternative text, font details, fill type, colors, pattern, and texture information—and writes the collected data to a formatted JSON file for external reporting or automation.
// Keywords: Aspose.Cells | C# | Excel shape export | shape properties JSON | font extraction | fill type | texture data | worksheet shape reporting | alternative text | rich text flag | automation
// Common Searches: Aspose.Cells export shape properties to JSON | read Excel shape font and fill with C# | how to get texture information from shapes in Aspose.Cells | serialize worksheet shapes to JSON | extract shape alignment and alternative text using Aspose.Cells
// Developer Intent: Retrieve all formatting and metadata of Excel shapes and serialize the information to JSON for downstream consumption.
// Use Cases: Generate an audit report that lists every shape’s text, styling, and fill details across a workbook. | Migrate shape formatting from a legacy workbook to a new template by reading the JSON and applying the settings programmatically. | Validate accessibility compliance by checking alternative text and rich‑text flags on all shapes.
// AI Prompts: Add shape rotation angle and Z‑order to the JSON output in the existing Aspose.Cells example. | Create a C# function that filters the exported JSON to include only shapes with a specific FillType or TextureType. | Refactor the program to stream the JSON payload directly to a REST API instead of writing a local file.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePropertiesExport
{
    // A C# utility that loads an Excel workbook using Aspose.Cells, walks through every worksheet, extracts each shape's properties—including name, type, text, alignment, rich‑text flag, alternative text, font details, fill type, colors, pattern, and texture information—and writes the collected data to a formatted JSON file for external reporting or automation.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "ShapeProperties.json";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // List to hold shape information
                var shapesInfo = new List<Dictionary<string, object>>();

                // Iterate through worksheets and their shapes
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        var shapeData = new Dictionary<string, object>
                        {
                            ["Worksheet"] = sheet.Name,
                            ["Name"] = shape.Name,
                            ["Type"] = shape.Type.ToString(),
                            ["Text"] = shape.Text,
                            ["IsRichText"] = shape.IsRichText,
                            ["AlternativeText"] = shape.AlternativeText,
                            ["TextHorizontalAlignment"] = shape.TextHorizontalAlignment.ToString(),
                            ["TextVerticalAlignment"] = shape.TextVerticalAlignment.ToString()
                        };

                        // Font information (if available)
                        if (shape.Font != null)
                        {
                            var fontInfo = new Dictionary<string, object>
                            {
                                ["Name"] = shape.Font.Name,
                                ["Size"] = shape.Font.Size,
                                ["Bold"] = shape.Font.IsBold,
                                ["Italic"] = shape.Font.IsItalic,
                                ["Underline"] = shape.Font.Underline,
                                ["Color"] = shape.Font.Color.ToArgb()
                            };
                            shapeData["Font"] = fontInfo;
                        }

                        // Fill information (if available)
                        if (shape.Fill != null)
                        {
                            var fillInfo = new Dictionary<string, object>
                            {
                                ["FillType"] = shape.Fill.FillType.ToString(),
                                ["Pattern"] = shape.Fill.Pattern.ToString()
                            };

                            // Attempt to read ForeColor and BackColor via reflection (may not exist in older versions)
                            try
                            {
                                var foreProp = shape.Fill.GetType().GetProperty("ForeColor");
                                if (foreProp != null)
                                {
                                    var foreVal = foreProp.GetValue(shape.Fill);
                                    if (foreVal is Color foreColor)
                                        fillInfo["ForeColor"] = foreColor.ToArgb();
                                }

                                var backProp = shape.Fill.GetType().GetProperty("BackColor");
                                if (backProp != null)
                                {
                                    var backVal = backProp.GetValue(shape.Fill);
                                    if (backVal is Color backColor)
                                        fillInfo["BackColor"] = backColor.ToArgb();
                                }
                            }
                            catch
                            {
                                // Ignore if properties are not supported
                            }

                            // Texture information (if a texture is applied)
                            try
                            {
                                var textureObj = shape.Fill.Texture;
                                if (textureObj != null)
                                {
                                    var textureInfo = new Dictionary<string, object>();

                                    var typeProp = textureObj.GetType().GetProperty("TextureType");
                                    if (typeProp != null)
                                    {
                                        var texType = typeProp.GetValue(textureObj);
                                        textureInfo["TextureType"] = texType?.ToString();
                                    }

                                    var imageProp = textureObj.GetType().GetProperty("Image");
                                    if (imageProp != null)
                                    {
                                        var img = imageProp.GetValue(textureObj);
                                        textureInfo["Image"] = img?.ToString();
                                    }

                                    if (textureInfo.Count > 0)
                                        fillInfo["Texture"] = textureInfo;
                                }
                            }
                            catch
                            {
                                // Ignore if texture details are not supported
                            }

                            shapeData["Fill"] = fillInfo;
                        }

                        shapesInfo.Add(shapeData);
                    }
                }

                // Serialize to JSON with indentation
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string jsonString = JsonSerializer.Serialize(shapesInfo, jsonOptions);

                // Write JSON to output file
                File.WriteAllText(outputPath, jsonString);

                Console.WriteLine($"Shape properties have been exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
