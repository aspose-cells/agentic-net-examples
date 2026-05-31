using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace TextBoxValidationExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("Template.xlsx");

            // Define the required tags that must be present in each TextBox
            // Example: tags are placeholders like {{Name}} and {{Date}}
            string[] requiredTags = new string[] { "{{Name}}", "{{Date}}" };

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of TextBox shapes on the current worksheet
                TextBoxCollection textBoxes = sheet.TextBoxes;

                // Loop through each TextBox
                for (int i = 0; i < textBoxes.Count; i++)
                {
                    TextBox tb = textBoxes[i];

                    // Retrieve the plain text of the TextBox
                    string text = tb.Text ?? string.Empty;

                    // Verify that all required tags are present
                    foreach (string tag in requiredTags)
                    {
                        if (!text.Contains(tag))
                        {
                            // If a required tag is missing, throw an informative exception
                            throw new InvalidOperationException(
                                $"TextBox '{tb.Name}' in worksheet '{sheet.Name}' is missing required tag '{tag}'.");
                        }
                    }

                    // Perform batch replacement for each tag (example replacements)
                    // In a real scenario, replacement values would come from your data source
                    text = text.Replace("{{Name}}", "John Doe");
                    text = text.Replace("{{Date}}", DateTime.Today.ToShortDateString());

                    // Apply the updated text back to the TextBox
                    tb.Text = text;
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("Result.xlsx");
        }
    }
}