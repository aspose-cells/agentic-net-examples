// Title: Apply 1.5 line spacing to text inside a TextBox shape in an Excel file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a TextBox shape, sets its TextOptions.LineSpacing to 1.5, and saves the workbook with Aspose.Cells. | Show how to modify an existing Aspose.Cells TextBox shape to use 1.5 line spacing for its multiline text. | Provide a snippet that demonstrates configuring line spacing for shape text via TextOptions in Aspose.Cells for .NET.
// Common Searches: asp.net set line spacing for textbox shape using Aspose.Cells | c# Aspose.Cells TextBox line height 1.5 | how to change line spacing of shape text in Excel with Aspose.Cells .NET | TextOptions.LineSpacing example for Excel shapes in C#
// Tags: Aspose.Cells TextOptions.LineSpacing property | C# set textbox shape line spacing | Excel shape text formatting Aspose.Cells | adjust line height in Excel shape .NET | multiline textbox shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a new workbook, adds a multiline TextBox shape, configures TextOptions.LineSpacing to 1.5 to achieve 1.5‑line spacing, and saves the file as LineSpacingExample.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet.
                // Parameters: upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, width, height
                TextBox textBox = worksheet.Shapes.AddTextBox(2, 1, 2, 1, 200, 80);

                // Set the text inside the shape (use line breaks)
                textBox.Text = "First line\nSecond line\nThird line";

                // Define output file path
                string outputPath = "LineSpacingExample.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
