// Title: Replace placeholder tags in Excel TextBox shapes using a dictionary with Aspose.Cells for .NET
// Common Searches: Aspose.Cells replace placeholder text in TextBox shape C# | C# dictionary based tag substitution in Excel TextBox using Aspose.Cells | Iterate worksheet shapes and update TextBox content with Aspose.Cells .NET | Batch replace <TAG_A> <TAG_B> <TAG_C> in an .xlsx file using Aspose.Cells | How to modify TextBox shapes in a workbook and save with Aspose.Cells
// Tags: Aspose.Cells replace text in TextBox shape | C# dictionary placeholder substitution Excel | Iterate worksheet shapes Aspose.Cells | Update TextBox content in .xlsx with Aspose.Cells | Save modified workbook after shape text change

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, loops through each worksheet and its shapes, identifies TextBox shapes, and replaces placeholder tags (<TAG_A>, <TAG_B>, <TAG_C>) with values from a C# dictionary. Updated text is written back to the TextBox, the output directory is created if needed, and the workbook is saved, with error handling for missing files and shape processing issues.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Placeholder tags and their replacement values
            var placeholderMap = new Dictionary<string, string>
            {
                { "<TAG_A>", "ReplacementA" },
                { "<TAG_B>", "ReplacementB" },
                { "<TAG_C>", "ReplacementC" }
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Process only TextBox shapes (fallback to string comparison for safety)
                        if (shape.Type.ToString() == "TextBox")
                        {
                            // Get the current text of the TextBox (handle possible null)
                            string text = shape.Text ?? string.Empty;

                            // Replace each placeholder with its corresponding value
                            foreach (KeyValuePair<string, string> kvp in placeholderMap)
                            {
                                text = text.Replace(kvp.Key, kvp.Value);
                            }

                            // Assign the updated text back to the TextBox
                            shape.Text = text;
                        }
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Error processing shape on sheet '{sheet.Name}': {exShape.Message}");
                    }
                }
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
