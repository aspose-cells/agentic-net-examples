// Title: Replace multiple placeholder tags in Excel TextBox shapes using Aspose.Cells for .NET (C#)
// Common Searches: C# replace <TAG_A> <TAG_B> placeholders in Excel TextBox with Aspose.Cells | Aspose.Cells iterate shapes and update TextBox text from a dictionary | How to perform bulk placeholder substitution in Excel shape text using Aspose.Cells .NET | Programmatically replace custom tags in Excel shape objects via Aspose.Cells | Using a dictionary to map placeholder tags to values in an Excel workbook with Aspose.Cells
// Tags: placeholder tag replacement in Excel shapes Aspose.Cells | dictionary based shape text update Aspose.Cells C# | iterate over worksheet shapes to modify TextBox Aspose.Cells | bulk replace custom tags in .xlsx using Aspose.Cells | programmatic TextBox content substitution Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an existing .xlsx file, defines a dictionary that maps placeholder tags (e.g., <TAG_A>, <TAG_B>, <TAG_C>) to replacement strings, iterates through every worksheet and each shape on the sheet, replaces any matching placeholders in TextBox shape text, ensures the output directory exists, and saves the modified workbook. It includes error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Map placeholders to replacement values
            var placeholderMap = new Dictionary<string, string>
            {
                { "<TAG_A>", "Replacement for A" },
                { "<TAG_B>", "Replacement for B" },
                { "<TAG_C>", "Replacement for C" }
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process shapes that contain text (e.g., TextBox)
                    if (!string.IsNullOrEmpty(shape.Text))
                    {
                        string text = shape.Text;

                        // Replace placeholders
                        foreach (var kvp in placeholderMap)
                        {
                            if (text.Contains(kvp.Key))
                            {
                                text = text.Replace(kvp.Key, kvp.Value);
                            }
                        }

                        // Update the shape text
                        shape.Text = text;
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
