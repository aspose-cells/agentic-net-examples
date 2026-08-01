// Title: Replace Multiple Placeholder Tags in an Aspose.Cells TextBox Using a Dictionary (C#)
// Description: Creates a workbook, adds a TextBox containing <TAG_A>, <TAG_B>, and <TAG_C>, maps each tag to a value with a Dictionary, and uses FontSettingCollection.Replace to substitute all placeholders before saving the file.
// Keywords: Aspose.Cells | C# replace placeholder TextBox | FontSettingCollection Replace | dictionary text replacement | Excel textbox placeholder | Aspose.Cells TextBox API | batch replace tags | shape text replace | Excel automation C# | Aspose.Cells example
// Common Searches: C# replace placeholders in Aspose.Cells TextBox | How to use FontSettingCollection.Replace with a dictionary | Aspose.Cells replace <TAG_A> <TAG_B> <TAG_C> | Batch replace tags in Excel shape using Aspose | Dictionary based text replacement Aspose.Cells
// Developer Intent: Replace several placeholder tags inside a worksheet TextBox with values supplied by a dictionary.
// Use Cases: Generate personalized greetings by inserting a user's name and location into a TextBox. | Insert dynamically generated code snippets into a documentation TextBox. | Create mail‑merge style worksheets where shape text is populated per record. | Automate report templates that contain variable placeholders in shapes. | Localize UI text within Excel shapes using a key‑value map.
// AI Prompts: Show a C# example that replaces <TAG_A>, <TAG_B>, and <TAG_C> in an Aspose.Cells TextBox using a Dictionary. | Explain how FontSettingCollection.Replace works for multiple placeholders in Aspose.Cells. | Provide code to iterate over a Dictionary and update TextBox text in an Excel workbook with Aspose.Cells. | Demonstrate batch replacement of placeholder tags in Excel shapes using Aspose.Cells API.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a TextBox containing <TAG_A>, <TAG_B>, and <TAG_C>, maps each tag to a value with a Dictionary, and uses FontSettingCollection.Replace to substitute all placeholders before saving the file.
class ReplacePlaceholdersInTextBox
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox to the worksheet
        int textBoxIndex = worksheet.TextBoxes.Add(2, 2, 200, 100);
        TextBox textBox = worksheet.TextBoxes[textBoxIndex];

        // Set initial text containing placeholder tags
        textBox.Text = "Hello <TAG_A>, welcome to <TAG_B>. Your code: <TAG_C>.";

        // Dictionary that maps each placeholder to its replacement value
        var placeholderMap = new Dictionary<string, string>
        {
            { "<TAG_A>", "Alice" },
            { "<TAG_B>", "Wonderland" },
            { "<TAG_C>", "XYZ123" }
        };

        // Get the FontSettingCollection (TextBody) of the textbox
        FontSettingCollection fontSettings = textBox.TextBody;

        // Replace each placeholder using FontSettingCollection.Replace(string, string)
        foreach (var entry in placeholderMap)
        {
            fontSettings.Replace(entry.Key, entry.Value);
        }

        // Save the workbook to a file
        workbook.Save("PlaceholderReplaced.xlsx");
    }
}
