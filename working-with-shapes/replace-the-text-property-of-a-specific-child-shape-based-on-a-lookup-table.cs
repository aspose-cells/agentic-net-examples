// Title: Update child shape Text values in an Excel worksheet using Aspose.Cells and a C# dictionary lookup
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, loops through every shape on a worksheet, and substitutes the shape's Text property using a provided Dictionary<string,string> mapping. | Demonstrate how to catch missing workbook files and individual shape processing errors while performing bulk text replacement on Excel shapes with Aspose.Cells. | Adapt the sample to modify only TextBox shapes, leave other shapes untouched, and save the result to a new workbook.
// Common Searches: how to replace shape text in Excel using Aspose.Cells C# | c# bulk update of textbox content in .xlsx based on key value pairs | iterate over worksheet shapes and change Text property with Aspose.Cells | Aspose.Cells replace specific child shape text using dictionary lookup | C# example for updating Excel shape text programmatically
// Tags: Aspose.Cells shape text replacement | C# dictionary lookup for Excel shapes | update child shape Text property Aspose.Cells | bulk modify Excel shape content .NET | iterate worksheet shapes Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an Excel workbook, iterates through all shapes on the first worksheet, replaces each shape's Text with a new value from a Dictionary lookup, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the lookup table: old text -> new text
            var lookup = new Dictionary<string, string>()
            {
                { "OldText1", "NewText1" },
                { "OldText2", "NewText2" },
                // add more mappings as needed
            };

            // Get the first worksheet (adjust index or name as required)
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Process shape text replacement
                    if (!string.IsNullOrEmpty(shape.Text) && lookup.TryGetValue(shape.Text, out string newText))
                    {
                        shape.Text = newText;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing shape: {ex.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
