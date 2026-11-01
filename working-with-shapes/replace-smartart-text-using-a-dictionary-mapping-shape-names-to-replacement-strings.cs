// Title: How to replace SmartArt shape text in an Excel workbook using a case‑insensitive dictionary with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing XLSX file with Aspose.Cells, iterates through all worksheets, and updates the Text property of SmartArt shapes whose Name matches entries in a case‑insensitive Dictionary<string,string>. | Generate a snippet that creates the output directory if it does not exist, applies dictionary‑driven text replacements to matching SmartArt shapes, and saves the modified workbook as a new file using Aspose.Cells. | Show how to safely handle a missing input file and catch exceptions while performing shape text updates in an Excel workbook with Aspose.Cells. | Demonstrate iterating the Shapes collection of each worksheet, checking for non‑empty shape.Text, and assigning a new value from a mapping dictionary.
// Common Searches: c# aspocells replace smartart text based on shape name dictionary | aspocells iterate worksheet shapes and modify text property | how to update specific SmartArt shapes in an existing Excel file using Aspose.Cells | save modified Excel workbook to a new file after shape text changes aspocells | case insensitive dictionary lookup for shape names in Aspose.Cells C#
// Tags: Aspose.Cells modify SmartArt shape content | C# dictionary based shape text replacement | iterate worksheet shapes Aspose.Cells | save workbook after shape changes Aspose.Cells | create output folder if missing C# | case-insensitive shape name lookup Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing XLSX workbook with Aspose.Cells, builds a case‑insensitive dictionary that maps SmartArt shape names to new text, walks through each worksheet and its Shapes collection, replaces the Text property of matching shapes, ensures the output directory exists, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Map shape names to the desired replacement text
            var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "SmartArt1", "Replacement Text for SmartArt 1" },
                { "SmartArt2", "Replacement Text for SmartArt 2" }
                // Add more mappings as needed
            };

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // If the shape's name matches a key in the dictionary, replace its text
                    if (replacements.TryGetValue(shape.Name, out string newText))
                    {
                        // For shapes that support text (e.g., TextBox, Rectangle, etc.)
                        if (!string.IsNullOrEmpty(shape.Text))
                        {
                            shape.Text = newText;
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
