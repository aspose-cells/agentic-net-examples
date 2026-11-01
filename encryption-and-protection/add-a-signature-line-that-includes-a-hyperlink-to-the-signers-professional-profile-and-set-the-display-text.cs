// Title: Insert a clickable signature line with a LinkedIn hyperlink into an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to insert a textbox signature line with a LinkedIn link using Aspose.Cells C# | Aspose.Cells C# add hyperlink to a shape in an .xlsx file | set display text for a hyperlink on a textbox in Excel with Aspose.Cells | create a signed Excel document with a clickable profile link using Aspose.Cells .NET
// Tags: Aspose.Cells insert textbox hyperlink | C# add hyperlink to shape Excel | Excel signature line textbox Aspose.Cells | Aspose.Cells set hyperlink address on shape | save workbook with linked signature line

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a TextBox shape named "SignatureLine" at row 10 column 2, sets its text to the signer's name, assigns a LinkedIn URL as a clickable hyperlink with a screen tip, centers the text, applies optional line formatting, and saves the file as SignatureLine.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Position and size for the signature line (row, column, width, height)
            int row = 10;
            int column = 2;
            int width = 200;
            int height = 50;

            // Add a TextBox shape that will serve as the signature line
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
            var shape = sheet.Shapes.AddTextBox(row, column, 0, 0, height, width);
            shape.Name = "SignatureLine";
            shape.Text = "John Doe";

            // Configure hyperlink (read‑only property, set its members)
            shape.Hyperlink.Address = "https://www.linkedin.com/in/johndoe";
            shape.Hyperlink.ScreenTip = "View John's professional profile";

            // Optional formatting
            shape.Line.Weight = 1;
            // DashStyle enum may not be available in some versions; omit if not supported
            // shape.Line.DashStyle = Aspose.Cells.Drawing.LineDashStyle.Solid;
            shape.TextHorizontalAlignment = TextAlignmentType.Center;
            shape.TextVerticalAlignment = TextAlignmentType.Center;

            // Determine full output path and ensure directory exists
            string outputPath = "SignatureLine.xlsx";
            string fullOutputPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(fullOutputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{fullOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
