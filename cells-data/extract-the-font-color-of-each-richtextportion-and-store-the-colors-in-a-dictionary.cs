using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ExtractRichTextColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape with some rich text
        Shape shape = sheet.Shapes.AddTextBox(2, 0, 2, 100, 200, 0);
        shape.Text = "Red Blue Green";

        // Apply different font colors to portions of the text
        shape.Characters(0, 3).Font.Color = Color.Red;      // "Red"
        shape.Characters(4, 4).Font.Color = Color.Blue;    // "Blue"
        shape.Characters(9, 5).Font.Color = Color.Green;   // "Green"

        // Retrieve all rich text formatting segments
        FontSetting[] richFormattings = shape.GetRichFormattings();

        // Dictionary to hold the start index of each portion and its font color
        Dictionary<int, Color> portionColors = new Dictionary<int, Color>();

        foreach (FontSetting setting in richFormattings)
        {
            // Use the start index as the key and store the associated font color
            portionColors[setting.StartIndex] = setting.Font.Color;
        }

        // Display the collected colors
        foreach (KeyValuePair<int, Color> kvp in portionColors)
        {
            Console.WriteLine($"StartIndex: {kvp.Key}, Color: {kvp.Value}");
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("RichTextColors.xlsx");
    }
}