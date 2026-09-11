// Title: Replace SmartArt shape text in an Excel workbook using a case‑insensitive dictionary with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, iterates all worksheets and their Shape objects, and substitutes each shape's Text property using a provided Dictionary<string,string> (case‑insensitive). | Add detailed console logging to the SmartArt text replacement routine so that the original text, the matched dictionary key, and the new text are recorded for every modified shape. | Encapsulate the SmartArt text replacement logic into a reusable method that accepts inputPath, outputPath, and a Dictionary<string,string>, then saves the updated workbook.
// Common Searches: how to change SmartArt shape text in an Excel file with Aspose.Cells C# | c# replace worksheet shape text using dictionary Aspose.Cells | bulk update SmartArt labels in .xlsx programmatically Aspose | case insensitive text replacement for Excel shapes using Aspose.Cells .NET
// Tags: Aspose.Cells SmartArt text substitution | C# iterate worksheet shapes | dictionary driven shape text update | case-insensitive text mapping Excel | save modified workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, walks through each worksheet's Shape collection, replaces any shape text that matches a key in a case‑insensitive dictionary, and saves the modified file to a new location.
class SmartArtTextReplacer
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Mapping of original text to replacement text
            var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Old Text 1", "New Text A" },
                { "Old Text 2", "New Text B" }
                // add more mappings as needed
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Replace shape text if it matches a key in the dictionary
                        string currentText = shape.Text;
                        if (!string.IsNullOrEmpty(currentText) && replacements.TryGetValue(currentText, out string newText))
                        {
                            shape.Text = newText;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log but continue processing other shapes
                        Console.WriteLine($"Failed to process shape on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
