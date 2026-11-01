// Title: Create and save an Excel workbook with a diagonal WordArt watermark while retaining all shape attributes using Aspose.Cells for .NET
// Common Searches: how to add a diagonal WordArt watermark to an Excel worksheet with Aspose.Cells C# | save Excel file as XLSX preserving WordArt shape attributes using Aspose.Cells | Aspose.Cells set TextEffect shape transparency and rotation before saving | place WordArt behind cells as a watermark in Aspose.Cells .NET | C# Aspose.Cells add TextEffect shape with custom font size and style
// Tags: add TextEffect shape Aspose.Cells | diagonal WordArt watermark Excel | retain shape properties when saving XLSX | configure shape transparency and rotation Aspose.Cells | send shape behind cells free-floating placement

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a TextEffect WordArt shape with the text "CONFIDENTIAL", applies 50% fill transparency, rotates it -45 degrees, sets free‑floating placement and a negative Z‑order to position it behind cells, and then saves the workbook as an XLSX file while preserving all shape attributes.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add WordArt (TextEffect) as a watermark
            // Parameters: preset effect, text, font name, font size, bold, italic,
            // upper left row/column, lower right row/column, width, height
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,   // preset effect
                "CONFIDENTIAL",                    // watermark text
                "Arial",                           // font name
                72,                                // font size (int)
                true,                              // bold
                false,                             // italic
                5,                                 // upper left row
                1,                                 // upper left column
                20,                                // lower right row
                10,                                // lower right column
                300,                               // width
                100);                              // height

            // Set fill transparency (50% transparent)
            wordArt.Fill.Transparency = 0.5;

            // Rotate the WordArt to appear as a diagonal watermark
            wordArt.RotationAngle = -45;

            // Send the shape behind the cells (so it behaves like a watermark)
            wordArt.Placement = PlacementType.FreeFloating;
            wordArt.ZOrderPosition = -1; // negative Z-order places it behind other objects

            // Optional: hide the shape's border line
            // (If the API version supports it, uncomment the line below)
            // wordArt.Line.IsVisible = false;

            // Optional: set a light gray fill color if the API version supports it
            // (If the API version supports it, uncomment the line below)
            // wordArt.Fill.SolidFillColor = Color.LightGray;

            // Ensure the output directory exists
            string outputPath = "WatermarkedWorkbook.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as XLSX preserving all shape attributes
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
