using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtTextReplacement
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string mappingPath = "Mapping.csv";
                const string outputPath = "Output.xlsx";

                // Verify template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the workbook (template containing SmartArt)
                Workbook workbook = new Workbook(templatePath);

                // Verify mapping file exists
                if (!File.Exists(mappingPath))
                {
                    Console.WriteLine($"Mapping file not found: {mappingPath}");
                    return;
                }

                // Load CSV mapping (shape name, new text) – format: Name,NewText
                var nameToText = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var line in File.ReadLines(mappingPath))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    // Split by first comma to allow commas inside new text
                    var parts = line.Split(new[] { ',' }, 2);
                    if (parts.Length == 2)
                    {
                        var shapeName = parts[0].Trim();
                        var newText = parts[1].Trim();
                        if (!nameToText.ContainsKey(shapeName))
                            nameToText.Add(shapeName, newText);
                    }
                }

                // Iterate through all worksheets and shapes
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Process only SmartArt shapes
                        if (shape.IsSmartArt)
                        {
                            // Convert SmartArt to grouped shapes
                            var smartArtResult = shape.GetResultOfSmartArt();

                            // Iterate each grouped shape inside the SmartArt
                            foreach (Shape innerShape in smartArtResult.GetGroupedShapes())
                            {
                                // Replace text if shape name matches mapping
                                if (nameToText.TryGetValue(innerShape.Name, out string replacement))
                                {
                                    innerShape.Text = replacement;
                                }
                            }
                        }
                    }
                }

                // Save the workbook with SmartArt update enabled
                var saveOptions = new OoxmlSaveOptions { UpdateSmartArt = true };
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}