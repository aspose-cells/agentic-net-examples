// Title: Replace SmartArt shape text in an Excel workbook using a dictionary with Aspose.Cells for .NET
// Common Searches: Aspose.Cells replace SmartArt alternative text using C# dictionary | C# code to update SmartArt shape names in Excel workbook with Aspose.Cells | Iterate worksheet shapes and set AlternativeText for SmartArt in .NET | Bulk replace SmartArt text in Excel using Aspose.Cells API | Dictionary driven SmartArt text replacement in Aspose.Cells for .NET
// Tags: smartart alternativetext update aspnet | shape text replacement aspocells | excel workbook smartart modification c# | mass smartart text update aspocells | c# worksheet shape enumeration aspocells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook (creating a placeholder if missing), defines a case‑insensitive dictionary that maps SmartArt shape names to new text, iterates through each worksheet and its shapes, and for every SmartArt shape whose name exists in the dictionary, assigns the replacement string to the shape's AlternativeText property before saving the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create an empty workbook if it does not.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a new workbook with a default worksheet
                workbook.Save(inputPath);   // optionally save the placeholder file
            }

            // Mapping of SmartArt shape names to replacement text
            var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "SmartArt1", "New Text for SmartArt 1" },
                { "SmartArt2", "New Text for SmartArt 2" }
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only SmartArt shapes that have a replacement entry
                    if (shape.IsSmartArt && replacements.ContainsKey(shape.Name))
                    {
                        // Aspose.Cells does not expose SmartArt node editing directly.
                        // As a fallback, set the alternative text of the shape to the desired value.
                        shape.AlternativeText = replacements[shape.Name];
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
