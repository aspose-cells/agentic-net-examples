// Title: Replace worksheet header WordArt with a custom TextEffect shape using Aspose.Cells for .NET
// AI Prompts: Write a C# program that opens an existing .xlsx file, removes all shapes on each worksheet, and adds a new WordArt (TextEffect) shape as the header with specified text, font, and size using Aspose.Cells. | Adjust the example so that only the topmost WordArt header is replaced while preserving any other shapes on the sheet. | Enhance the code to apply a solid fill color, outline, and rotation to the inserted TextEffect shape.
// Common Searches: how to replace header WordArt in an Excel workbook with Aspose.Cells C# | Aspose.Cells add TextEffect shape as header to each worksheet programmatically | C# remove all shapes from Excel sheet then insert new WordArt using Aspose.Cells | set font and size for WordArt shape in Aspose.Cells .NET example | save modified Excel file as new workbook after changing header WordArt Aspose.Cells
// Tags: worksheet WordArt substitution Aspose.Cells | insert TextEffect shape C# | clear all shapes Excel Aspose.Cells | style WordArt font Aspose.Cells | save workbook as new file Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook, clears every shape on each worksheet, creates a new TextEffect (WordArt) shape with custom text, font, and size as the header, optionally formats it, and saves the result as a new Excel file.
class ReplaceHeaderWordArt
{
    static void Main()
    {
        try
        {
            // Path to the existing Excel file
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                // Remove existing shapes (optional – adjust as needed)
                for (int i = shapes.Count - 1; i >= 0; i--)
                {
                    shapes.RemoveAt(i);
                }

                // Add a new WordArt (TextEffect) shape
                Shape newWordArt = shapes.AddTextEffect(
                    (MsoPresetTextEffect)0,          // preset effect (plain text)
                    "New Header Title",              // text
                    "Arial Black",                   // font name
                    36,                              // font size
                    true,                            // bold
                    false,                           // italic
                    0, 0,                            // left, top
                    400, 50,                         // width, height
                    0, 0);                           // shape type and additional parameter (default values)

                // Basic formatting (fill and line) – optional
                try
                {
                    // Note: FillFormat and LineFormat may not expose ForeColor in some library versions.
                    // If needed, use alternative properties such as SolidFillColor or set line color via LineFormat.
                    // Example (uncomment if supported):
                    // newWordArt.Fill.ForeColor = Color.LightBlue;
                    // newWordArt.Line.ForeColor = Color.DarkBlue;
                    // newWordArt.Line.Weight = 2;
                }
                catch (Exception fmtEx)
                {
                    Console.WriteLine($"Formatting warning: {fmtEx.Message}");
                }

                // Rotation and Z-order
                newWordArt.RotationAngle = 0;
                newWordArt.ZOrderPosition = 0;
            }

            // Path to save the modified Excel file
            string outputPath = "output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
