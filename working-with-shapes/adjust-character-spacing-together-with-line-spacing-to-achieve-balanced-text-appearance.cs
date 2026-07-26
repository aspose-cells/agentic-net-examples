// Title: Adjust character and line spacing in a text box shape with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a text box, increase character spacing via TextOptions.Spacing, and define paragraph line spacing, space before, and space after in points to achieve a balanced layout before saving the file.
// Keywords: Aspose.Cells .NET | text box character spacing | line spacing points | TextOptions.Spacing | TextParagraph line space | SpaceBefore SpaceAfter | Excel shape formatting | adjust paragraph spacing | balanced text appearance | C# Aspose.Cells example
// Common Searches: Aspose.Cells increase character spacing | set line spacing for shape text in Excel | C# adjust paragraph spacing in text box | how to use TextOptions.Spacing Aspose.Cells | configure SpaceBefore and SpaceAfter Aspose.Cells
// Developer Intent: Apply both character and paragraph spacing settings to a text box shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Enhance readability of multi‑line text inside a shape for reports | Create consistently spaced headings within Excel dashboards | Prepare printable worksheets where text layout must meet design guidelines | Programmatically modify existing workbooks to standardize shape text formatting
// AI Prompts: Generate C# code that sets TextOptions.Spacing and paragraph LineSpace, SpaceBefore, SpaceAfter for a text box in Aspose.Cells. | Explain how to retrieve and modify spacing properties of existing text box shapes in an Excel file using Aspose.Cells. | Show an example that applies point‑based line spacing and uniform character spacing to all paragraphs inside a shape.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AdjustSpacingExample
{
    // Shows how to create a workbook, add a text box, increase character spacing via TextOptions.Spacing, and define paragraph line spacing, space before, and space after in points to achieve a balanced layout before saving the file.
    public class AdjustSpacingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, width, height
                Shape textBox = sheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 150);
                textBox.Text = "First line of text\nSecond line of text\nThird line";

                // Adjust character spacing for the entire text run
                TextOptions textOpts = textBox.TextOptions;
                textOpts.Spacing = 1.5; // increase spacing between characters

                // Access each paragraph to set line spacing and surrounding space
                TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;
                foreach (TextParagraph para in paragraphs)
                {
                    // Use points as the unit for line spacing
                    para.LineSpaceSizeType = LineSpaceSizeType.Points;
                    para.LineSpace = 8; // 8 points line spacing

                    // Add space before and after each paragraph for balanced appearance
                    para.SpaceBeforeSizeType = LineSpaceSizeType.Points;
                    para.SpaceAfterSizeType = LineSpaceSizeType.Points;
                    para.SpaceBefore = 2; // 2 points before
                    para.SpaceAfter = 2;  // 2 points after
                }

                // Save the workbook with the adjusted spacing
                string outputPath = "AdjustedSpacing.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AdjustSpacingDemo.Run();
        }
    }
}
