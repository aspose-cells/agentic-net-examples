// Title: C# – Replace Text in a Specific Excel Shape Using a Lookup Dictionary with Aspose.Cells
// Description: Loads an Excel workbook, finds a shape named "MyTextBox" on the first worksheet, and replaces its TextBody content using key‑value pairs from a Dictionary before saving the file.
// Keywords: Aspose.Cells replace shape text | C# Excel shape text replacement | Aspose.Cells TextBody Replace | lookup dictionary Excel textbox | update Excel shape content .NET | Aspose.Cells shape manipulation
// Common Searches: replace text in a specific shape Aspose.Cells C# | use dictionary to update Excel textbox with Aspose | Aspose.Cells change placeholder text in shape | iterate worksheet shapes and modify TextBody | C# replace shape text based on lookup table
// Developer Intent: Replace the Text property of a targeted child shape using a lookup dictionary.
// Use Cases: Swap placeholder strings in template workbooks with real data before distribution. | Apply language translation dictionaries to shape labels for localization. | Insert calculated values such as totals or dates into specific shape text boxes.
// AI Prompts: Generate C# code that iterates through worksheet shapes and replaces text using a Dictionary<string,string> with Aspose.Cells. | Show how to perform case‑insensitive replacements in a shape's TextBody using Aspose.Cells for .NET. | Provide an example that logs each text replacement applied to a shape's TextBody during processing.

using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Loads an Excel workbook, finds a shape named "MyTextBox" on the first worksheet, and replaces its TextBody content using key‑value pairs from a Dictionary before saving the file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that contains the shape
        Worksheet worksheet = workbook.Worksheets[0];

        // Lookup table: key = text to find, value = text to replace with
        var lookup = new Dictionary<string, string>
        {
            { "OldValue1", "NewValue1" },
            { "Placeholder", "ActualData" },
            { "ABC", "XYZ" }
        };

        // Locate the specific child shape (by name, index, or any other criteria)
        Shape targetShape = null;
        foreach (Shape shape in worksheet.Shapes)
        {
            // Example: identify shape by its Name property
            if (shape.Name == "MyTextBox")
            {
                targetShape = shape;
                break;
            }
        }

        if (targetShape != null)
        {
            // Use the shape's TextBody (FontSettingCollection) to perform replacements
            FontSettingCollection textBody = targetShape.TextBody;

            foreach (var pair in lookup)
            {
                // Replace all occurrences of the old text with the new text
                textBody.Replace(pair.Key, pair.Value);
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
