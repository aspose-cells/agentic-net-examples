using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPartialTextColor
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text box shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset, upper left offset, width, height
            Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 300, 100);
            textBox.Text = "Partial color change example";

            // Define the range of characters to recolor (e.g., "color")
            int startIndex = textBox.Text.IndexOf("color");
            int length = "color".Length;

            // Get the FontSetting for the specified range
            FontSetting fontSetting = textBox.Characters(startIndex, length);

            // Set a custom RGB color (e.g., orange)
            fontSetting.Font.Color = Color.FromArgb(255, 165, 0); // RGB(255,165,0)

            // Save the workbook
            workbook.Save("PartialTextColor.xlsx");
        }
    }
}