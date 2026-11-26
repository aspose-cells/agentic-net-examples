// Title: Add a WordArt shape with a custom Comic Sans MS font at 48 pt using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add WordArt shape with custom font size 48 points | How to change the font family of a WordArt shape in an Excel workbook with Aspose.Cells | Set Comic Sans MS font for WordArt text using Aspose.Cells .NET | Apply StyleFlag to format the text body of a WordArt shape in C# | Create WordArt with a preset style and custom font in Aspose.Cells for .NET
// Tags: add WordArt shape Aspose.Cells C# | custom font family WordArt Aspose.Cells | WordArt font size 48 points | StyleFlag format WordArt text body | preset WordArt style Excel Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace WordArtExample
{
    // This example shows how to create a new workbook, insert a WordArt shape with a preset style, assign the text "Custom WordArt", define a custom style using the Comic Sans MS font at 48 pt, apply the font name and size to the entire WordArt text via a StyleFlag, and save the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt shape with a preset style
            // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height (pixel), width (pixel)
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Custom WordArt",
                0,      // topRow
                0,      // top offset in pixels
                0,      // leftColumn
                0,      // left offset in pixels
                100,    // height in pixels
                400);   // width in pixels

            // Access the text body of the WordArt shape
            FontSettingCollection textBody = wordArt.TextBody;

            // Ensure the text is set (AddWordArt already sets it, but we can overwrite if needed)
            textBody.Text = "Custom WordArt";

            // Create a style to define the custom font family and size
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.Name = "Comic Sans MS";   // Custom font family
            customStyle.Font.Size = 48;                // Font size in points

            // Define which style attributes should be applied
            StyleFlag flag = new StyleFlag();
            flag.FontName = true;
            flag.FontSize = true;

            // Apply the style to the entire text of the WordArt
            int textLength = textBody.Text.Length;
            if (textLength > 0)
            {
                textBody.Format(0, textLength, customStyle.Font, flag);
            }

            // Save the workbook to a file
            string outputPath = "WordArtCustomFont.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
