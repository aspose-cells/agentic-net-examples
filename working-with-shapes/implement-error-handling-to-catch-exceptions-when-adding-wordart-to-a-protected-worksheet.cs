// Title: Add WordArt to a protected worksheet with exception handling using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add WordArt to a locked worksheet with try‑catch | how to handle AddTextEffect exception on a protected Excel sheet in .NET | protect Excel worksheet then insert text effect using Aspose.Cells C# example | catch errors when inserting shapes into a protected workbook with Aspose.Cells | Aspose.Cells C# sample for adding WordArt after worksheet protection
// Tags: add wordart to protected worksheet Aspose.Cells | exception handling for AddTextEffect C# | protect worksheet before inserting shapes Aspose.Cells | c# workbook save after shape addition Aspose.Cells | handling shape insertion errors Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, protects the first worksheet, attempts to add a WordArt text effect inside a try‑catch block, logs any exception, ensures the output directory exists, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet (all protection types)
            sheet.Protect(ProtectionType.All);

            try
            {
                // Add WordArt (text effect) to the protected worksheet
                // Parameters: preset, text, font name, font size, bold, italic,
                // upper left row, upper left column, top offset, left offset,
                // height, width
                sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "Protected WordArt",
                    "Arial",
                    20,
                    false,
                    false,
                    1,
                    1,
                    0,
                    0,
                    100,   // height
                    300);  // width
            }
            catch (Exception ex)
            {
                // Display any exception that occurs while adding WordArt
                Console.WriteLine("Error adding WordArt to a protected worksheet: " + ex.Message);
            }

            // Save the workbook
            string outputPath = "ProtectedWorksheetWithWordArt.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
