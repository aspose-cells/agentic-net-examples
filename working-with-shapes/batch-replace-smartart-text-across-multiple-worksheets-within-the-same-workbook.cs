// Title: How to batch replace SmartArt shape text in all worksheets of an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells replace text in SmartArt shapes across multiple worksheets | batch update SmartArt text in every sheet of an Excel file with Aspose.Cells | iterate through workbook worksheets to modify shape text using Aspose.Cells .NET | dictionary driven text replacement for Excel SmartArt using Aspose.Cells | programmatically change SmartArt labels in all sheets with Aspose.Cells C#
// Tags: Aspose.Cells batch SmartArt text replacement | C# iterate workbook worksheets shapes | dictionary based shape text update Excel | replace SmartArt labels across multiple sheets | process Excel shapes with Aspose.Cells .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, defines a dictionary of old‑to‑new text pairs, loops through each worksheet and each shape, replaces matching substrings in SmartArt shape text, and saves the modified workbook to a new file.
class SmartArtBatchReplace
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsx";
        const string outputPath = "OutputWorkbook.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Define the text replacements: key = text to find, value = replacement text
        var replacements = new Dictionary<string, string>
        {
            { "OldText1", "NewText1" },
            { "OldText2", "NewText2" }
            // add more pairs as needed
        };

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Process shapes that contain text (e.g., text boxes)
                    if (!string.IsNullOrEmpty(shape.Text))
                    {
                        foreach (var pair in replacements)
                        {
                            if (shape.Text.Contains(pair.Key))
                            {
                                shape.Text = shape.Text.Replace(pair.Key, pair.Value);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log shape‑specific errors and continue processing other shapes
                    Console.WriteLine($"Error processing shape on sheet '{sheet.Name}': {ex.Message}");
                }
            }
        }

        try
        {
            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
