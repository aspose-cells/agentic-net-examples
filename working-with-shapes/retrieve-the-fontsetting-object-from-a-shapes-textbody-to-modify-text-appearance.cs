// Title: Aspose.Cells C# – Retrieve and Edit FontSetting in a Shape TextBody
// Description: Demonstrates how to add a rectangle shape to a workbook, assign text, obtain the FontSettingCollection via the shape's TextBody, and apply character‑level font changes such as name, size, color, bold and italic before saving the file.
// Keywords: Aspose.Cells FontSetting | C# shape TextBody | character level formatting | FontSettingCollection example | modify shape text font | Aspose.Cells .NET tutorial
// Common Searches: Aspose.Cells change font of specific characters in a shape | How to get FontSetting from shape TextBody C# | Apply color and style to a range of characters in Aspose.Cells shape | Set bold font for first character of shape text Aspose.Cells
// Developer Intent: Access the FontSetting objects of a shape’s TextBody to style individual characters programmatically.
// Use Cases: Set the first character of a shape’s text to Arial, 16 pt, blue, and bold. | Apply red italic formatting to characters 8‑13 within the shape’s text. | Iterate over FontSettingCollection after assigning text to a shape to apply mixed styles before saving the workbook.
// AI Prompts: Generate C# code that retrieves a FontSetting for a given character index from a shape’s TextBody and changes its font to Times New Roman, 12 pt, green. | Create a loop that formats characters 5‑10 in a shape’s TextBody to be underlined with a yellow background using Aspose.Cells. | Explain step‑by‑step how to access FontSettingCollection after setting shape.TextBody.Text and modify multiple character styles in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsFontSettingDemo
{
    // Demonstrates how to add a rectangle shape to a workbook, assign text, obtain the FontSettingCollection via the shape's TextBody, and apply character‑level font changes such as name, size, color, bold and italic before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 200, 100);

                // Set the shape's text via the TextBody collection (populates FontSettingCollection)
                shape.TextBody.Text = "Aspose.Cells FontSetting Example";

                // Retrieve the FontSettingCollection from the shape's TextBody
                FontSettingCollection fontSettings = shape.TextBody;

                // Modify the whole text if needed (already set above)
                // fontSettings.Text = "Aspose.Cells FontSetting Example";

                // Retrieve a specific FontSetting by character index (e.g., index 0 for the first character)
                // The collection contains a FontSetting for each character after the text is set
                FontSetting firstCharSetting = fontSettings[0];

                // Modify the font appearance of the first character
                firstCharSetting.Font.Name = "Arial";
                firstCharSetting.Font.Size = 16;
                firstCharSetting.Font.Color = Color.Blue;
                firstCharSetting.Font.IsBold = true;

                // Apply formatting to a range of characters (indices 8 to 13)
                for (int i = 8; i <= 13 && i < fontSettings.Count; i++)
                {
                    FontSetting setting = fontSettings[i];
                    setting.Font.Color = Color.Red;
                    setting.Font.IsItalic = true;
                }

                // Save the workbook
                workbook.Save("FontSettingDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
