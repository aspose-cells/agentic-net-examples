// Title: Replace SmartArt shape text in an Excel workbook using a CSV mapping with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, reads a CSV file containing shape‑name and new‑label pairs, builds a case‑insensitive dictionary, and updates each SmartArt shape whose current label matches a key. | Create a method that iterates through all worksheets and their Shapes collection, checks the Shape.Text property, substitutes it with the mapped value, and saves the modified workbook to a new file. | Add try‑catch blocks around shape processing to log warnings for shapes that cannot be updated while allowing the rest of the SmartArt update operation to continue.
// Common Searches: C# Aspose.Cells replace SmartArt text from CSV file | bulk update Excel SmartArt shape labels using a mapping file | read shape name to text mapping CSV and apply to Excel workbook with Aspose.Cells | iterate over worksheet shapes and change text in Aspose.Cells .NET | handle errors when updating SmartArt shapes in Aspose.Cells
// Tags: Aspose.Cells update SmartArt shape labels | C# CSV mapping to modify Excel shapes | bulk update SmartArt shapes Aspose.Cells | shape processing error handling Aspose.Cells | iterate worksheet shapes with Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, reads a CSV file of shape‑name to new‑text mappings into a case‑insensitive dictionary, iterates through every worksheet and each shape, updates the shape's text when it matches a dictionary key, logs warnings for any shape that cannot be processed, and saves the result as a new .xlsx file.
class SmartArtTextReplacer
{
    static void Main()
    {
        // Paths to the Excel file and the CSV mapping file
        string excelPath = "input.xlsx";
        string csvPath = "mapping.csv";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input files exist
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"Excel file not found: {excelPath}");
            if (!File.Exists(csvPath))
                throw new FileNotFoundException($"CSV mapping file not found: {csvPath}");

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Build a dictionary from the CSV file: shapeName -> newText
            var mapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in File.ReadAllLines(csvPath))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Expected format: ShapeName,NewValue
                var parts = line.Split(new[] { ',' }, 2);
                if (parts.Length == 2)
                {
                    string shapeName = parts[0].Trim();
                    string newValue = parts[1].Trim();
                    if (!mapping.ContainsKey(shapeName))
                        mapping.Add(shapeName, newValue);
                }
            }

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Attempt to replace the shape's text if it matches a key in the mapping
                        string currentText = shape.Text?.Trim();
                        if (!string.IsNullOrEmpty(currentText) && mapping.TryGetValue(currentText, out string replacement))
                        {
                            shape.Text = replacement;
                        }
                    }
                    catch (Exception exShape)
                    {
                        // Log shape‑level errors but continue processing other shapes
                        Console.WriteLine($"Warning: Could not process shape '{shape.Name}'. {exShape.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
