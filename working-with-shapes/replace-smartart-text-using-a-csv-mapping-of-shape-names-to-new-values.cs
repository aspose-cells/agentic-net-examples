// Title: Replace SmartArt shape text in an Excel workbook using a CSV mapping with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace SmartArt shape text from CSV file | how to update Excel SmartArt shapes using a mapping dictionary in .NET | read shape name and new text from CSV and set shape.Text with Aspose.Cells | iterate all worksheets and shapes to change SmartArt text in C# | bulk replace SmartArt text in Excel workbook using external CSV mapping
// Tags: Aspose.Cells SmartArt text replacement from CSV | C# shape.Text property update in Excel | load CSV mapping dictionary for workbook automation | iterate worksheet shapes Aspose.Cells | replace SmartArt shapes using name mapping .NET | save modified workbook as Xlsx with Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an input.xlsx workbook, reads a mapping.csv file that pairs SmartArt shape names with new text, builds a case‑insensitive dictionary, iterates every worksheet and its shapes, and for each shape whose Name matches a key in the dictionary sets the Shape.Text property to the mapped value. The modified workbook is then saved as output.xlsx.
class SmartArtTextReplacer
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string mappingPath = "mapping.csv";
            const string outputPath = "output.xlsx";

            // Verify input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Workbook file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Load CSV mapping (shape name, new text) into a dictionary
            var mapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (File.Exists(mappingPath))
            {
                try
                {
                    foreach (var line in File.ReadAllLines(mappingPath))
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var parts = line.Split(new[] { ',' }, 2);
                        if (parts.Length == 2)
                        {
                            var shapeName = parts[0].Trim();
                            var newText = parts[1].Trim();
                            mapping[shapeName] = newText;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to read mapping file: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Warning: Mapping file \"{mappingPath}\" not found. No replacements will be applied.");
            }

            // Iterate through all worksheets and replace shape texts based on the mapping
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape == null)
                        continue;

                    if (mapping.TryGetValue(shape.Name, out string replacement))
                    {
                        try
                        {
                            // For shapes that support text, set the Text property
                            shape.Text = replacement;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to set text for shape \"{shape.Name}\": {ex.Message}");
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
