// Title: Validate required placeholder tags in every TextBox shape before batch replacement with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook using Aspose.Cells, iterates through all TextBox shapes, checks each for a predefined list of placeholders (e.g., {{Name}}, {{Date}}, {{Amount}}), logs any missing tags, and aborts the operation if validation fails. | Generate a C# method that performs placeholder substitution inside TextBox shapes only after confirming that every shape contains all required tags, then saves the updated workbook. | Create a C# example that collects validation errors for missing placeholders in Excel TextBox shapes and throws an InvalidOperationException when any required tag is absent.
// Common Searches: c# Aspose.Cells verify required placeholders in Excel textbox shapes before replace | how to abort batch placeholder substitution when a textbox shape is missing a tag in Aspose.Cells | log missing {{Name}} tag in Excel textbox using Aspose.Cells .NET | validate all TextBox shapes for {{Date}} and {{Amount}} placeholders with Aspose.Cells | batch replace placeholder values in Excel textbox after validation Aspose.Cells
// Tags: Aspose.Cells validate textbox placeholders | C# substitute Excel textbox placeholder values | detect missing shape tags Aspose.Cells | Excel shape tag verification .NET | placeholder substitution workflow Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, scans every worksheet's shapes, validates that each TextBox contains the required tags {{Name}}, {{Date}} and {{Amount}}, reports any missing tags and aborts with an exception, and if validation passes performs a batch substitution of those placeholders with actual values before saving the workbook.
class TextBoxTagValidator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Required tags that must be present in every TextBox shape
            List<string> requiredTags = new List<string> { "{{Name}}", "{{Date}}", "{{Amount}}" };

            // Collect validation errors
            List<string> validationErrors = new List<string>();

            // Iterate through all worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only TextBox shapes (use AutoShapeType for compatibility)
                    if (shape.Type == AutoShapeType.TextBox)
                    {
                        string text = shape.Text;

                        foreach (string tag in requiredTags)
                        {
                            if (!text.Contains(tag))
                            {
                                validationErrors.Add(
                                    $"Worksheet \"{sheet.Name}\", Shape \"{shape.Name}\": missing tag {tag}");
                            }
                        }
                    }
                }
            }

            // Abort if validation fails
            if (validationErrors.Count > 0)
            {
                Console.WriteLine("Validation failed. Missing required tags:");
                foreach (string error in validationErrors)
                {
                    Console.WriteLine(error);
                }
                throw new InvalidOperationException("One or more TextBox shapes are missing required tags.");
            }

            // ----- Batch Replacement (executed only after successful validation) -----
            Dictionary<string, string> replacementMap = new Dictionary<string, string>
            {
                { "{{Name}}", "John Doe" },
                { "{{Date}}", DateTime.Today.ToString("yyyy-MM-dd") },
                { "{{Amount}}", "1234.56" }
            };

            // Perform replacement inside each TextBox shape
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.Type == AutoShapeType.TextBox)
                    {
                        string updatedText = shape.Text;

                        foreach (var kvp in replacementMap)
                        {
                            updatedText = updatedText.Replace(kvp.Key, kvp.Value);
                        }

                        shape.Text = updatedText;
                    }
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
