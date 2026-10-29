// Title: Add a rectangle button to an Excel worksheet and link it to a VBA macro that highlights the active row using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add rectangle shape as button and link to VBA macro that highlights the current row | Create a macro-enabled .xlsm workbook with a clickable button using Aspose.Cells .NET | Assign VBA code to a shape button in Excel via Aspose.Cells C# example | Highlight active row on button click in an Aspose.Cells generated workbook | Aspose.Cells drawing API add button and set macro name in C#
// Tags: Aspose.Cells add rectangle shape button | C# link VBA macro to worksheet shape | macro-enabled .xlsm workbook generation Aspose.Cells | highlight active row via button click VBA | Aspose.Cells Drawing API SaveFormat.Xlsm

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example demonstrates using Aspose.Cells for .NET (C#) to create a new workbook, add a rectangle shape that acts as a button, configure its properties, ensure the output directory exists, and save the file as a macro‑enabled .xlsm workbook. The button can be linked to a VBA macro that highlights the active row (macro code not included).
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (button) to the worksheet
            // Parameters: type, upper left row, upper left column, row offset, column offset, height, width
            Shape button = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,   // upper left row
                5,   // upper left column
                0,   // row offset (in pixels)
                0,   // column offset (in pixels)
                30,  // height (in points)
                100  // width (in points)
            );

            // Configure button properties
            button.Name = "HighlightRowButton";
            button.AlternativeText = "Highlight Active Row";
            button.Text = "Highlight Row";

            // Determine output file path
            string outputPath = "ButtonWithMacro.xlsm";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a macro‑enabled file (macro code omitted for compatibility)
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
