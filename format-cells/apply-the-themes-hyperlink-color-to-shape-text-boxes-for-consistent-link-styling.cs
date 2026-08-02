using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ApplyHyperlinkThemeColor
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a TextBox shape at cell B2 (row=1, column=1) with size 200x50 pixels
            TextBox textBox = (TextBox)sheet.Shapes.AddTextBox(1, 1, 0, 0, 200, 50);

            // Set the display text of the TextBox
            textBox.Text = "Visit Aspose";

            // Add a hyperlink to the entire shape
            textBox.AddHyperlink("https://www.aspose.com");

            // Retrieve the theme's hyperlink color
            Color hyperlinkThemeColor = workbook.GetThemeColor(ThemeColorType.Hyperlink);

            // Apply the theme hyperlink color to the TextBox text
            textBox.Font.Color = hyperlinkThemeColor;

            // Save the workbook
            string outputPath = "HyperlinkStyledTextbox.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}