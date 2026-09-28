// Title: Replace multiple placeholder tags in an Excel TextBox using a dictionary with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells to replace placeholders like <TAG_A>, <TAG_B>, and <TAG_C> in a TextBox shape using a Dictionary<string,string>. | Extend the example to loop through all TextBox shapes on a worksheet and apply dictionary‑based placeholder replacement. | Create a case‑insensitive placeholder substitution routine for TextBox objects in an Excel workbook using Aspose.Cells.
// Common Searches: how to replace placeholder tags in an Excel textbox using Aspose.Cells C# | dictionary based text replacement for shapes in Aspose.Cells workbook | update multiple textboxes in an Excel sheet with values from a C# dictionary Aspose.Cells | case insensitive placeholder substitution in Excel shapes using Aspose.Cells .NET
// Tags: textbox placeholder replacement Aspose.Cells C# | dictionary driven text update Excel shape | iterate worksheet shapes Aspose.Cells | case insensitive tag substitution Aspose.Cells | bulk text replacement Excel textbox .NET

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, finds the first TextBox shape on the first worksheet, replaces the placeholders <TAG_A>, <TAG_B>, and <TAG_C> with values from a dictionary, writes the updated text back to the TextBox, and saves the workbook.
class ReplaceTextBoxPlaceholders
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the first worksheet (adjust as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Find the first TextBox shape in the worksheet
        // You can also locate by shape name or index if known
        TextBox textbox = null;
        foreach (Shape shape in sheet.Shapes)
        {
            if (shape is TextBox)
            {
                textbox = (TextBox)shape;
                break; // use the first textbox found
            }
        }

        if (textbox == null)
        {
            Console.WriteLine("No TextBox found in the worksheet.");
            return;
        }

        // Dictionary mapping placeholders to replacement values
        var placeholderMap = new Dictionary<string, string>
        {
            { "<TAG_A>", "Value for A" },
            { "<TAG_B>", "Value for B" },
            { "<TAG_C>", "Value for C" }
        };

        // Get the current text of the TextBox
        string text = textbox.Text;

        // Replace each placeholder with its corresponding value
        foreach (var kvp in placeholderMap)
        {
            if (!string.IsNullOrEmpty(kvp.Key))
            {
                text = text.Replace(kvp.Key, kvp.Value);
            }
        }

        // Assign the modified text back to the TextBox
        textbox.Text = text;

        // Save the workbook (replace with your desired output path)
        workbook.Save("Output.xlsx");
    }
}
