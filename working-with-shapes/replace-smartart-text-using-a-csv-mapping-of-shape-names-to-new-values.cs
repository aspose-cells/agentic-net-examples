// Title: Replace SmartArt Text in Excel via CSV Mapping – Aspose.Cells for .NET
// Description: Load a template workbook, read a CSV that maps SmartArt shape names to new values, convert each SmartArt to a GroupShape, update matching inner shapes, and save the file with UpdateSmartArt enabled using Aspose.Cells C# API.
// Keywords: Aspose.Cells SmartArt replace text | C# update Excel SmartArt from CSV | Excel SmartArt text mapping | GroupShape inner shapes Aspose.Cells | UpdateSmartArt option | .NET Excel automation | batch SmartArt editing
// Common Searches: how to change SmartArt text in Excel with Aspose.Cells | C# replace SmartArt shape labels using a CSV file | Aspose.Cells read CSV and update SmartArt | access inner shapes of SmartArt in .NET | save Excel workbook with updated SmartArt
// Developer Intent: Programmatically modify the text of SmartArt shapes in an Excel workbook based on a CSV file that maps shape names to replacement strings.
// Use Cases: Populate organization‑chart nodes from an employee list stored in CSV. | Refresh process‑flow diagram labels across many worksheets with data from an external source. | Generate personalized reports where SmartArt captions are filled automatically before distribution.
// AI Prompts: Generate C# code using Aspose.Cells to read a CSV of shape names and replace corresponding SmartArt text, ensuring the workbook is saved with UpdateSmartArt=true. | Explain how to retrieve inner shapes of a SmartArt object in Aspose.Cells when the GroupShapes collection is not publicly exposed. | Create robust error handling for missing template or CSV files while performing SmartArt text replacement with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

namespace AsposeCellsSmartArtReplace
{
    // Load a template workbook, read a CSV that maps SmartArt shape names to new values, convert each SmartArt to a GroupShape, update matching inner shapes, and save the file with UpdateSmartArt enabled using Aspose.Cells C# API.
    public class SmartArtReplacer
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("SmartArt replacement completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Paths to the template workbook, CSV mapping file and output workbook
            string templatePath = "template.xlsx";
            string csvPath = "mapping.csv";
            string outputPath = "output.xlsx";

            // Verify required files exist
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template workbook not found: {templatePath}");
            if (!File.Exists(csvPath))
                throw new FileNotFoundException($"CSV mapping file not found: {csvPath}");

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(templatePath);

            // Read CSV mapping (shape name, new text) into a dictionary
            var textMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string line in File.ReadAllLines(csvPath))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue; // skip empty lines

                string[] parts = line.Split(new[] { ',' }, 2);
                if (parts.Length == 2)
                {
                    string shapeName = parts[0].Trim();
                    string newText = parts[1].Trim();
                    textMap[shapeName] = newText;
                }
            }

            // Iterate through all worksheets and shapes
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (!shape.IsSmartArt)
                        continue;

                    try
                    {
                        // Convert SmartArt to a GroupShape
                        GroupShape groupShape = shape.GetResultOfSmartArt();

                        // Use reflection to obtain inner shapes (compatible with older Aspose.Cells versions)
                        PropertyInfo innerProp = groupShape.GetType().GetProperty("GroupShapes");
                        if (innerProp != null)
                        {
                            var innerCollection = innerProp.GetValue(groupShape) as IEnumerable<Shape>;
                            if (innerCollection != null)
                            {
                                foreach (Shape innerShape in innerCollection)
                                {
                                    if (textMap.TryGetValue(innerShape.Name, out string replacement))
                                    {
                                        innerShape.Text = replacement;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to process SmartArt shape '{shape.Name}'. {ex.Message}");
                    }
                }
            }

            // Save the workbook with UpdateSmartArt enabled (lifecycle rule: save)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                UpdateSmartArt = true
            };
            workbook.Save(outputPath, saveOptions);
        }
    }
}
