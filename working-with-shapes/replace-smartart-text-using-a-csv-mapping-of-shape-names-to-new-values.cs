// Title: Update SmartArt shape text in an Excel workbook using a CSV mapping with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, reads a CSV file of shape names and new texts into a case‑insensitive dictionary, and updates the Text property of any SmartArt shape whose current text matches a key. | Create a robust CSV‑parsing method that handles commas inside values, builds the mapping dictionary, and applies it to modify SmartArt shapes across all worksheets. | Add comprehensive error handling that validates the presence of the workbook and CSV files, creates the output folder when missing, and saves the modified workbook to a target path.
// Common Searches: how to change SmartArt text in Excel using Aspose.Cells and a CSV file in C# | replace Excel shape names with new values from a mapping file using Aspose.Cells .NET | C# example for iterating worksheet shapes and updating SmartArt text based on CSV data
// Tags: Aspose.Cells SmartArt text replacement via CSV | Excel shape name mapping with Aspose.Cells | iterate worksheet shapes Aspose.Cells | C# update SmartArt text in .xlsx | load CSV into dictionary for Excel processing

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample validates that both the source workbook and the CSV mapping file exist, loads the workbook with Aspose.Cells, reads the CSV into a case‑insensitive dictionary of shape names to replacement texts, iterates through every worksheet and its shapes, updates the Text property of each SmartArt shape whose current text matches a dictionary key, ensures the output directory is present, and saves the modified workbook to the specified location.
class Program
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
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Verify mapping file exists
            if (!File.Exists(mappingPath))
                throw new FileNotFoundException($"Mapping CSV not found: {mappingPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Load the CSV mapping of shape names to new text values
            Dictionary<string, string> mapping = LoadCsvMapping(mappingPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only SmartArt shapes (or any shape with text)
                    if (shape.IsSmartArt)
                    {
                        // Use the shape's text as the key for replacement
                        string currentText = shape.Text;

                        if (mapping.TryGetValue(currentText, out string newText))
                        {
                            shape.Text = newText;
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to read a CSV file where each line is: ShapeName,NewValue
    static Dictionary<string, string> LoadCsvMapping(string csvPath)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (string line in File.ReadLines(csvPath))
        {
            // Skip empty or whitespace-only lines
            if (string.IsNullOrWhiteSpace(line))
                continue;

            // Split on the first comma to allow commas inside the new value
            int commaIndex = line.IndexOf(',');
            if (commaIndex <= 0)
                continue; // Invalid line format

            string key = line.Substring(0, commaIndex).Trim();
            string value = line.Substring(commaIndex + 1).Trim();

            if (!string.IsNullOrEmpty(key))
                dict[key] = value;
        }

        return dict;
    }
}
