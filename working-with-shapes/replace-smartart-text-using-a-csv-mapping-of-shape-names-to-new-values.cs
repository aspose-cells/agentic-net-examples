// Title: Replace SmartArt Text in Excel Using CSV Mapping – Aspose.Cells for .NET
// Description: This example demonstrates how to load an Excel workbook, read a CSV file that maps SmartArt shape names to new text values, locate SmartArt objects, convert them to grouped shapes, replace matching shape text, and save the file with UpdateSmartArt enabled using Aspose.Cells for .NET.
// Keywords: Aspose.Cells SmartArt text replacement | C# update SmartArt from CSV | Excel SmartArt shape mapping | replace SmartArt labels programmatically | UpdateSmartArt option Aspose | batch edit SmartArt Excel | CSV driven Excel diagram text change
// Common Searches: how to change SmartArt text in Excel with Aspose.Cells | C# replace SmartArt shape names using CSV | update Excel SmartArt programmatically .NET | Aspose.Cells map shape name to new text | replace SmartArt labels in bulk Excel
// Developer Intent: Programmatically replace the text of SmartArt shapes in an Excel workbook based on a CSV file that maps shape names to new values.
// Use Cases: Localize SmartArt diagrams by applying translated strings from a CSV file. | Populate dynamic data into SmartArt charts for automated report generation. | Apply corporate branding updates across all SmartArt objects in multiple worksheets.
// AI Prompts: Generate C# code with Aspose.Cells that reads a CSV of shape names and replaces matching SmartArt text, ensuring UpdateSmartArt is set when saving. | Explain the role of GetResultOfSmartArt().GetGroupedShapes() for accessing individual SmartArt elements in Aspose.Cells. | Suggest robust error‑handling strategies for missing files, unmapped shape names, or empty CSV entries when updating SmartArt.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example demonstrates how to load an Excel workbook, read a CSV file that maps SmartArt shape names to new text values, locate SmartArt objects, convert them to grouped shapes, replace matching shape text, and save the file with UpdateSmartArt enabled using Aspose.Cells for .NET.
    class SmartArtReplaceDemo
    {
        public static void Run()
        {
            try
            {
                const string templatePath = "template.xlsx";
                const string mappingPath = "mapping.csv";
                const string outputPath = "output.xlsx";

                // Verify required files exist
                if (!File.Exists(templatePath))
                    throw new FileNotFoundException($"Template file not found: {templatePath}");
                if (!File.Exists(mappingPath))
                    throw new FileNotFoundException($"Mapping file not found: {mappingPath}");

                // Load the source workbook
                Workbook workbook = new Workbook(templatePath);

                // Load CSV mapping (shape name, new text)
                var mapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var line in File.ReadAllLines(mappingPath))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var parts = line.Split(new[] { ',' }, 2);
                    if (parts.Length == 2)
                    {
                        var name = parts[0].Trim();
                        var newText = parts[1].Trim();
                        mapping[name] = newText;
                    }
                }

                // Iterate through all worksheets and shapes
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        // Process only SmartArt shapes
                        if (shape.IsSmartArt)
                        {
                            // Convert SmartArt to grouped shapes
                            var groupedShapes = shape.GetResultOfSmartArt().GetGroupedShapes();

                            // Replace text of each grouped shape based on the CSV mapping
                            foreach (Shape smartShape in groupedShapes)
                            {
                                if (!string.IsNullOrEmpty(smartShape.Name) &&
                                    mapping.TryGetValue(smartShape.Name, out string replacement))
                                {
                                    smartShape.Text = replacement;
                                }
                            }
                        }
                    }
                }

                // Save the workbook with SmartArt update enabled
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true
                };
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SmartArtReplaceDemo.Run();
        }
    }
}
