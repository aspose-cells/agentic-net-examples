// Title: Validate required placeholder tags in Excel TextBox shapes before batch replacement with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells verify that every TextBox shape contains {{CustomerName}} placeholder before replacement | How to detect missing required tags in Excel TextBox objects using Aspose.Cells .NET API | Batch replace placeholders in worksheet TextBox shapes after tag validation with Aspose.Cells | Throw exception for absent tags in Excel TextBox shapes when processing workbooks in C#
// Tags: Aspose.Cells validate TextBox placeholders | Excel TextBox tag verification C# | batch replace TextBox placeholders Aspose.Cells | required tag check worksheet shapes .NET | exception missing TextBox tag Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through each worksheet and its shapes, and ensures every TextBox shape contains a predefined set of placeholder tags (e.g., {{CustomerName}}, {{OrderDate}}, {{TotalAmount}}). If a tag is missing, an InvalidOperationException is thrown with details. After successful validation, the code performs a batch replacement of those placeholders with supplied values and saves the updated workbook.
class TextBoxTagValidator
{
    // List of required tags that must be present in every TextBox shape
    private static readonly List<string> RequiredTags = new List<string>
    {
        "{{CustomerName}}",
        "{{OrderDate}}",
        "{{TotalAmount}}"
    };

    // Sample replacement values for batch replacement
    private static readonly Dictionary<string, string> ReplacementValues = new Dictionary<string, string>
    {
        { "{{CustomerName}}", "Acme Corp" },
        { "{{OrderDate}}", DateTime.Today.ToShortDateString() },
        { "{{TotalAmount}}", "$1,234.56" }
    };

    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Validate TextBox shapes for required tags
            ValidateTextBoxTags(workbook);

            // Perform batch replacement after validation succeeds
            BatchReplaceTextBoxPlaceholders(workbook);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
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
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    private static void ValidateTextBoxTags(Workbook workbook)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get all shapes on the worksheet
            ShapeCollection shapes = sheet.Shapes;

            foreach (Shape shape in shapes)
            {
                // Process only TextBox shapes
                if (shape.AutoShapeType == AutoShapeType.TextBox)
                {
                    // Retrieve the text content of the TextBox (null‑safe)
                    string text = shape.Text ?? string.Empty;

                    // Check each required tag
                    foreach (string tag in RequiredTags)
                    {
                        if (!text.Contains(tag))
                        {
                            // Throw an exception with details about the missing tag
                            throw new InvalidOperationException(
                                $"Missing required tag '{tag}' in TextBox on sheet '{sheet.Name}'.");
                        }
                    }
                }
            }
        }
    }

    private static void BatchReplaceTextBoxPlaceholders(Workbook workbook)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            ShapeCollection shapes = sheet.Shapes;

            foreach (Shape shape in shapes)
            {
                // Process only TextBox shapes
                if (shape.AutoShapeType == AutoShapeType.TextBox)
                {
                    string text = shape.Text ?? string.Empty;

                    // Replace each placeholder with its corresponding value
                    foreach (var kvp in ReplacementValues)
                    {
                        if (text.Contains(kvp.Key))
                        {
                            text = text.Replace(kvp.Key, kvp.Value);
                        }
                    }

                    // Update the TextBox with the new text
                    shape.Text = text;
                }
            }
        }
    }
}
