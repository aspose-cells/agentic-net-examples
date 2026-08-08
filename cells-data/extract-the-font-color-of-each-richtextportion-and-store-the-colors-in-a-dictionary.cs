// Title: C# – Extract Font Colors from RichTextPortion Objects in a Shape with Aspose.Cells
// Description: This example creates a workbook, adds a rectangle shape with rich‑text, assigns different Font.Color values to each text portion, uses GetRichFormattings to read the colors, stores them in a Dictionary<int, Color>, prints the mapping, and saves the file.
// Keywords: Aspose.Cells GetRichFormattings | RichTextPortion font color C# | extract shape text colors Aspose.Cells | dictionary of font colors .NET | C# Aspose.Cells rich text formatting | shape text color extraction
// Common Searches: how to read font colors of RichTextPortion in Aspose.Cells | C# get RichFormattings colors from a shape | store Aspose.Cells text portion colors in a dictionary | Aspose.Cells extract rich text colors .NET | retrieve shape text colors using Aspose.Cells
// Developer Intent: Read the Font.Color of every RichTextPortion in a shape and map each color to its portion index.
// Use Cases: Create a legend that reflects the exact colors used in shape labels. | Validate that specific keywords in a diagram are highlighted with the correct colors before publishing. | Drive conditional formatting or automation based on the colors extracted from shape text.
// AI Prompts: Generate C# code that iterates over a shape's RichFormattings and fills a Dictionary<int, Color> with each portion's Font.Color using Aspose.Cells. | Show how to replace a particular font color in the extracted dictionary with a new color and apply the change back to the workbook. | Provide step‑by‑step instructions to display the portion index and its color in a console application.

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a rectangle shape with rich‑text, assigns different Font.Color values to each text portion, uses GetRichFormattings to read the colors, stores them in a Dictionary<int, Color>, prints the mapping, and saves the file.
class ExtractRichTextPortionColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape with rich text
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 0);
        shape.Text = "Red Blue Green";

        // Apply different font colors to portions of the text
        shape.Characters(0, 3).Font.Color = Color.Red;      // "Red"
        shape.Characters(4, 4).Font.Color = Color.Blue;    // "Blue"
        shape.Characters(9, 5).Font.Color = Color.Green;   // "Green"

        // Retrieve rich text formatting information
        FontSetting[] richFormattings = shape.GetRichFormattings();

        // Store the font color of each portion in a dictionary
        // Key: portion index, Value: font color
        Dictionary<int, Color> portionColors = new Dictionary<int, Color>();
        for (int i = 0; i < richFormattings.Length; i++)
        {
            portionColors[i] = richFormattings[i].Font.Color;
        }

        // Display the extracted colors
        foreach (var kvp in portionColors)
        {
            Console.WriteLine($"Portion {kvp.Key}: Color = {kvp.Value}");
        }

        // Save the workbook (optional, just to complete the lifecycle)
        workbook.Save("RichTextColors.xlsx");
    }
}
