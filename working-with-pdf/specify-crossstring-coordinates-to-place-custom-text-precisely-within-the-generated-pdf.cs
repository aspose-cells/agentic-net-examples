using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Add some data to the sheet
            worksheet.Cells["A1"].PutValue("Sample Data");

            // Add a TextBox shape. The parameters are:
            // upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
            // This overload returns the created TextBox object directly.
            TextBox textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 200, 50);

            // Specify the exact position in pixels relative to the worksheet's top‑left corner
            textBox.X = 150; // 150 pixels from the left edge
            textBox.Y = 300; // 300 pixels from the top edge

            // Set the custom text to be displayed
            textBox.Text = "Precise custom text";

            // Ensure the textbox size adapts to the text (optional)
            textBox.TextBody.TextAlignment.AutoSize = true;

            // Define output file name
            string outputPath = "PreciseText.pdf";

            // Save the workbook as PDF; the textbox will appear at the defined coordinates
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}