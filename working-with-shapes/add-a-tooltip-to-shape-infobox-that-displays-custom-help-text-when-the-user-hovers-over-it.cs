using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddTooltipToInfoBox
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox shape (InfoBox) to the worksheet
            // Parameters: upper left row, upper left column, upper left offset (pixels),
            //             lower right row, lower right column, lower right offset (pixels)
            TextBox infoBox = worksheet.Shapes.AddTextBox(2, 2, 0, 5, 5, 0);

            // Set the displayed text inside the textbox
            infoBox.Text = "InfoBox";

            // Set the tooltip text that appears when the user hovers over the shape
            infoBox.AlternativeText = "Custom help text displayed as a tooltip for the InfoBox.";

            // Define output file path
            string outputPath = "InfoBoxWithTooltip.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}