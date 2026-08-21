// Title: Get and Edit FontSetting of a Shape’s TextBody in Aspose.Cells for .NET
// Description: Demonstrates how to access a shape’s TextBody, retrieve its FontSettingCollection, modify individual FontSetting objects (e.g., first character or a character range), and save the workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# shape FontSetting | TextBody font formatting | modify shape text font | shape Characters method | .NET spreadsheet API | font color bold italic Aspose
// Common Searches: Aspose.Cells change font of specific characters in a shape | retrieve FontSettingCollection from a rectangle shape | format part of shape text using Characters method | C# Aspose.Cells set bold color for first character | how to edit shape text font in Aspose.Cells .NET
// Developer Intent: Extract FontSetting objects from a shape’s TextBody and adjust their font attributes programmatically.
// Use Cases: Set the first character of a shape’s text to Calibri, 16 pt, blue, and bold. | Apply dark‑red color and italic style to the word "Cells" (characters 7‑12) within the shape. | Iterate through all FontSetting entries to apply conditional formatting such as alternating colors.
// AI Prompts: Generate C# code with Aspose.Cells that retrieves a shape’s FontSettingCollection and makes the first character blue, bold, and 16 pt. | Show how to use the Characters method to italicize and color a specific substring of a shape’s text in Aspose.Cells for .NET. | Provide an example that loops over each FontSetting in a shape’s TextBody to apply custom formatting based on runtime logic.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to access a shape’s TextBody, retrieve its FontSettingCollection, modify individual FontSetting objects (e.g., first character or a character range), and save the workbook using Aspose.Cells for C#.
class RetrieveFontSettingFromShape
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offsetX, offsetY, width, height
        Shape shape = sheet.Shapes.AddRectangle(1, 1, 0, 0, 150, 80);

        // Set the shape's text
        shape.Text = "Aspose Cells";

        // Retrieve the FontSettingCollection from the shape's TextBody
        FontSettingCollection textBody = shape.TextBody;

        // Access a specific FontSetting (e.g., the first character)
        // The collection contains a FontSetting for each character in the text
        FontSetting firstCharSetting = textBody[0];

        // Modify the font appearance of the selected characters
        firstCharSetting.Font.Name = "Calibri";
        firstCharSetting.Font.Size = 16;
        firstCharSetting.Font.Color = Color.Blue;
        firstCharSetting.Font.IsBold = true;

        // Optionally, modify another range using the Characters method
        // Here we format characters 7 to 12 ("Cells")
        FontSetting rangeSetting = shape.Characters(7, 5);
        rangeSetting.Font.Color = Color.DarkRed;
        rangeSetting.Font.IsItalic = true;

        // Save the workbook
        workbook.Save("RetrieveFontSettingFromShape.xlsx");
    }
}
