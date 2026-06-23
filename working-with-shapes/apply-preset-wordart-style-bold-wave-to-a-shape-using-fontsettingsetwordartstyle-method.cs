using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class ApplyBoldWaveWordArt
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text box shape to the worksheet
            // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
            Shape textBox = worksheet.Shapes.AddTextBox(2, 10, 2, 10, 100, 300);
            textBox.Text = "Bold Wave Example";

            // Apply a preset WordArt style to the text using FontSetting
            // FontSetting(startIndex, length, worksheets)
            FontSetting fontSetting = new FontSetting(0, textBox.Text.Length, workbook.Worksheets);
            // Choose any preset style; here we use WordArtStyle3 as an example
            fontSetting.SetWordArtStyle(PresetWordArtStyle.WordArtStyle3);

            // Enhance the appearance to achieve a "Bold Wave" effect
            // Set the text to bold and apply the Wave1 shape
            TextEffectFormat textEffect = textBox.TextEffect;
            textEffect.FontBold = true; // Bold
            textEffect.PresetShape = MsoPresetTextEffectShape.Wave1; // Wave

            // Define output file path
            string outputPath = "BoldWaveWordArt.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}