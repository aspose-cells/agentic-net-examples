// Title: C# – Replace Multiple Placeholder Tags in an Aspose.Cells TextBox via Dictionary Mapping
// Description: Creates a workbook, adds a TextBox, sets text with <TAG_A>, <TAG_B>, <TAG_C>, then iterates a Dictionary<string,string> to substitute each placeholder using FontSettingCollection.Replace, and saves the file as PlaceholderReplaced.xlsx.
// Keywords: Aspose.Cells TextBox placeholder replacement | C# dictionary text substitution Aspose.Cells | FontSettingCollection Replace example | Aspose.Cells replace tags in worksheet | dynamic TextBox content Aspose.Cells
// Common Searches: Aspose.Cells replace placeholders in TextBox C# | Dictionary based tag substitution Aspose.Cells | How to use FontSettingCollection.Replace with TextBox | C# replace <TAG_A> <TAG_B> <TAG_C> in Excel workbook | Aspose.Cells TextBox dynamic text example
// Developer Intent: Swap placeholder tags inside a TextBox with actual values by looping through a dictionary and calling FontSettingCollection.Replace.
// Use Cases: Populate a report template TextBox with calculated metrics. | Generate a product list by mapping SKU codes to names. | Localize UI strings in a worksheet by replacing language tags.
// AI Prompts: Show C# code that uses Aspose.Cells to replace several placeholder tags in a TextBox using a Dictionary<string,string>. | Explain how to iterate over a dictionary and apply FontSettingCollection.Replace for each entry in a TextBox. | Demonstrate saving the workbook after performing placeholder substitutions with Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a TextBox, sets text with <TAG_A>, <TAG_B>, <TAG_C>, then iterates a Dictionary<string,string> to substitute each placeholder using FontSettingCollection.Replace, and saves the file as PlaceholderReplaced.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a TextBox to the worksheet
        int textBoxIndex = worksheet.TextBoxes.Add(2, 2, 200, 100);
        TextBox textBox = worksheet.TextBoxes[textBoxIndex];

        // Set initial text containing placeholder tags
        textBox.Text = "Values: <TAG_A>, <TAG_B>, and <TAG_C>.";

        // Dictionary that maps placeholders to their replacement values
        var placeholderMap = new Dictionary<string, string>
        {
            { "<TAG_A>", "Apple" },
            { "<TAG_B>", "Banana" },
            { "<TAG_C>", "Cherry" }
        };

        // Get the FontSettingCollection (text body) of the TextBox
        FontSettingCollection fontSettings = textBox.TextBody;

        // Replace each placeholder using the Replace(string, string) method (rule)
        foreach (var kvp in placeholderMap)
        {
            fontSettings.Replace(kvp.Key, kvp.Value);
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("PlaceholderReplaced.xlsx");
    }
}
