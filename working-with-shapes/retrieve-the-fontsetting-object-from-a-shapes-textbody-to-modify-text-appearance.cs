// Title: Retrieve and modify a FontSetting from a shape's TextBody using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get FontSetting from a shape's TextBody | change font style of text inside a shape with Aspose.Cells .NET | apply color and underline to a substring of shape text in Aspose.Cells | modify bold italic size of shape text via FontSettingCollection Aspose.Cells | format specific character range in a shape's TextBody using Aspose.Cells C#
// Tags: shape TextBody FontSettingCollection manipulation | modify shape text font attributes Aspose.Cells | apply style to subrange of shape text Aspose.Cells | C# Aspose.Cells shape font formatting | retrieve first FontSetting from shape TextBody

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace FontSettingFromShapeDemo
{
    // The example creates a workbook, adds a rectangle shape, accesses its TextBody (a FontSettingCollection), retrieves the first FontSetting, and updates font properties such as bold, italic, size, color, and name. It also demonstrates using TextBody.Format with a Style and StyleFlag to color and underline a specific character range, then saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output Excel file
            string outputPath = "ShapeFontSettingDemo.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, lower right row, lower right column, width (pixels), height (pixels)
                Shape rectangle = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 200, 100);
                rectangle.Name = "DemoRectangle";

                // 3. Access the TextBody (FontSettingCollection) of the shape
                FontSettingCollection textBody = rectangle.TextBody;

                // 4. Set the initial text for the shape
                textBody.Text = "Aspose.Cells FontSetting Demo";

                // 5. Retrieve the first FontSetting object from the collection
                // After setting text, the collection contains at least one FontSetting representing the whole text run
                FontSetting firstFontSetting = textBody[0];

                // 6. Modify the appearance of the text via the FontSetting object
                firstFontSetting.Font.IsBold = true;
                firstFontSetting.Font.IsItalic = true;
                firstFontSetting.Font.Size = 18;
                firstFontSetting.Font.Color = Color.DarkBlue;
                firstFontSetting.Font.Name = "Calibri";

                // 7. Optionally, demonstrate using the Format method to change a sub‑range
                // Change characters 8‑14 ("Cells") to red and underline
                Style style = workbook.CreateStyle();
                style.Font.Color = Color.Red;
                style.Font.Underline = FontUnderlineType.Single; // Correct enum usage

                StyleFlag flag = new StyleFlag();
                flag.FontColor = true;
                flag.FontUnderline = true;

                // Apply the style to the specified range (start index, length)
                textBody.Format(8, 5, style.Font, flag);

                // 8. Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
