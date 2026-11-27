// Title: How to add padding to a textbox shape in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set left and right margins of a textbox shape | adjust padding inside Excel textbox with TextBoxOptions in .NET | disable auto‑size and define internal margins for a shape textbox using Aspose.Cells | set top and bottom margin points for a textbox shape in a generated XLSX file | C# example of configuring textbox padding in Aspose.Cells workbook
// Tags: Aspose.Cells TextBoxOptions margin settings | C# set textbox shape padding points | Excel shape internal margins Aspose.Cells | textbox autosize off Aspose.Cells | save workbook with padded textbox Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// The example creates a workbook, inserts a textbox shape, turns off its auto‑size feature, applies left, right, top, and bottom internal margins in points via TextBoxOptions, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet.
            // Parameters: upper left row, upper left column, row offset (pixels), column offset (pixels), width (points), height (points)
            Shape textBox = worksheet.Shapes.AddTextBox(1, 1, 0, 0, 200, 100);
            textBox.Text = "This is a sample text inside the textbox. Adjusting margins adds padding around the text.";

            // Ensure the textbox does not auto‑size so that the margins are respected
            textBox.TextBody.TextAlignment.AutoSize = false;

            // Set internal margins (padding) in points. 1 point = 1/72 inch.
            textBox.TextBoxOptions.LeftMarginPt   = 12; // left padding
            textBox.TextBoxOptions.RightMarginPt  = 12; // right padding
            textBox.TextBoxOptions.TopMarginPt    = 8;  // top padding
            textBox.TextBoxOptions.BottomMarginPt = 8;  // bottom padding

            // Save the workbook to a file.
            string outputPath = "MarginTextBox.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
