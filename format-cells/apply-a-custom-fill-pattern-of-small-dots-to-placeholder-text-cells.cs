// Title: Apply Dotted‑Grid Fill Pattern to an AutoShape Placeholder in Aspose.Cells for .NET
// Description: Creates a new workbook, inserts a rectangle AutoShape as a placeholder cell, sets its text, applies a dotted‑grid pattern fill (light‑gray dots on white), optionally styles the text, and saves the file as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells pattern fill | dotted grid AutoShape | placeholder shape Excel | C# fill pattern Aspose | custom shape fill color | Excel template placeholder | Aspose.Cells .NET example
// Common Searches: how to add dotted grid fill to AutoShape Aspose.Cells | set pattern fill for placeholder shape in C# | Aspose.Cells custom fill pattern example | change foreground/background colors of shape fill | format placeholder cell with pattern in Excel
// Developer Intent: Add a rectangle AutoShape, assign placeholder text, and apply a small‑dot (dotted‑grid) fill pattern.
// Use Cases: Design Excel templates where input fields are highlighted with a dotted‑grid background. | Generate reports that separate sections using patterned rectangle shapes. | Create form‑style worksheets with placeholder shapes that visually indicate empty cells.
// AI Prompts: Write C# code with Aspose.Cells to insert a rectangle AutoShape, set placeholder text, and apply a dotted‑grid fill pattern. | Show how to apply the same pattern fill to multiple AutoShapes and customize their foreground and background colors. | Explain step‑by‑step how to change an AutoShape's FillType to Pattern and configure FillPattern.DottedGrid in Aspose.Cells for .NET.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExample
{
    // Creates a new workbook, inserts a rectangle AutoShape as a placeholder cell, sets its text, applies a dotted‑grid pattern fill (light‑gray dots on white), optionally styles the text, and saves the file as XLSX using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle AutoShape that will act as a placeholder text cell
                // Parameters: type, upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
                Shape placeholder = sheet.Shapes.AddAutoShape(
                    AutoShapeType.Rectangle,
                    5,          // Upper‑left row index
                    5,          // Upper‑left column index
                    0,          // Row offset (in pixels)
                    0,          // Column offset (in pixels)
                    50,         // Height in points
                    200         // Width in points
                );

                // Set placeholder text
                placeholder.Text = "Placeholder Text";

                // Apply a small‑dot (dotted grid) fill pattern to the shape
                placeholder.Fill.FillType = FillType.Pattern;
                placeholder.Fill.PatternFill.Pattern = FillPattern.DottedGrid;
                placeholder.Fill.PatternFill.ForegroundColor = Color.LightGray; // Dot color
                placeholder.Fill.PatternFill.BackgroundColor = Color.White;     // Background color

                // Optional: format the placeholder text
                FontSetting fontSetting = placeholder.Characters(0, placeholder.Text.Length);
                TextOptions textOpts = fontSetting.TextOptions;
                textOpts.Name = "Arial";
                textOpts.Size = 12;
                textOpts.IsBold = true;

                // Determine output file path
                string outputFile = "PlaceholderPattern.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
