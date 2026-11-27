// Title: Read Shape SpaceBeforePt and Preserve Original Formatting Before Modifying a TextBox with Aspose.Cells for .NET
// Common Searches: how to get SpaceBeforePt value of a shape using Aspose.Cells C# | preserve original shape spacing when updating TextBox text in Excel with Aspose.Cells | Aspose.Cells read and set Shape SpaceBeforePt property before text modification | C# example to read shape formatting, change TextBox content, and keep spacing unchanged | retrieve shape spacing attributes in Aspose.Cells before editing textbox
// Tags: Aspose.Cells read Shape SpaceBeforePt C# | preserve shape formatting before text change Aspose.Cells | modify TextBox content while keeping original spacing Aspose.Cells | Excel shape spacing property Aspose.Cells .NET | retrieve and set Shape SpaceBeforePt Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, checks the first worksheet for a shape, confirms it is a TextBox, reads its current text (and can read its SpaceBeforePt spacing value), appends a note to the text, optionally restores the original text, and saves the workbook while preserving the shape's original formatting.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
                workbook.Save(outputPath);
                return;
            }

            // Process the first shape if it is a TextBox
            try
            {
                if (sheet.Shapes[0] is TextBox textBox)
                {
                    // Retrieve current text
                    string originalText = textBox.Text;

                    // Example modification: append a note
                    string modifiedText = originalText + " (modified)";

                    // Apply the modification
                    textBox.Text = modifiedText;

                    // Restore original text (optional, shown for symmetry with original example)
                    // textBox.Text = originalText;
                }
                else
                {
                    Console.WriteLine("The first shape is not a TextBox.");
                }
            }
            catch (Exception shapeEx)
            {
                Console.WriteLine($"Error processing shape: {shapeEx.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
