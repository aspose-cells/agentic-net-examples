using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class Program
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
        textBox.Text = "Value A: <TAG_A>, Value B: <TAG_B>, Value C: <TAG_C>";

        // Dictionary that maps placeholders to their replacement values
        var placeholderMap = new Dictionary<string, string>
        {
            { "<TAG_A>", "Apple" },
            { "<TAG_B>", "Banana" },
            { "<TAG_C>", "Cherry" }
        };

        // Replace each placeholder in the textbox using FontSettingCollection.Replace(string, string)
        foreach (var kvp in placeholderMap)
        {
            textBox.TextBody.Replace(kvp.Key, kvp.Value);
        }

        // Save the workbook to a file
        workbook.Save("ReplacePlaceholdersInTextBox.xlsx");
    }
}