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
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will contain rich text
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 0);
        shape.Text = "RedBlueGreen";

        // Apply different font colors to portions of the text
        shape.Characters(0, 3).Font.Color = Color.Red;      // "Red"
        shape.Characters(3, 4).Font.Color = Color.Blue;     // "Blue"
        shape.Characters(7, 5).Font.Color = Color.Green;    // "Green"

        // Retrieve all rich‑text formatting objects from the shape
        FontSetting[] richFormattings = shape.GetRichFormattings();

        // Dictionary to store the start index of each portion and its font color
        Dictionary<int, Color> portionColors = new Dictionary<int, Color>();

        // Extract the color from each FontSetting and add it to the dictionary
        foreach (FontSetting setting in richFormattings)
        {
            portionColors[setting.StartIndex] = setting.Font.Color;
        }

        // Display the extracted colors
        foreach (KeyValuePair<int, Color> entry in portionColors)
        {
            Console.WriteLine($"StartIndex: {entry.Key}, Color: {entry.Value}");
        }

        // Save the workbook (optional, demonstrates lifecycle rule usage)
        workbook.Save("RichTextColors.xlsx");
    }
}