// Title: How to lock a WordArt watermark shape in an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# lock WordArt shape so users cannot move or edit it | prevent selection of watermark shape in an Excel workbook with Aspose.Cells | make WordArt watermark read‑only in generated .xlsx using Aspose.Cells | C# set IsLocked true on text effect shape in Aspose.Cells | add non‑editable WordArt watermark to Excel sheet programmatically
// Tags: Aspose.Cells shape lock API | WordArt text effect watermark Excel | C# free floating shape placement | read‑only watermark in generated .xlsx | disable shape editing Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a WordArt shape with the text "CONFIDENTIAL" as a watermark on the first worksheet, sets the shape's placement to FreeFloating, locks the shape using the IsLocked property to stop selection, movement, editing, and resizing, and saves the result as WatermarkedLocked.xlsx.
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

            // Add a WordArt shape that will serve as a watermark
            // Parameters: preset text effect, text, font name, font size, bold, italic,
            // upper left row, upper left column, lower right row, lower right column,
            // width, height
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset WordArt effect
                "CONFIDENTIAL",
                "Arial",
                72,
                false,
                false,
                0,   // upper row
                0,   // upper column
                30,  // lower row
                10,  // lower column
                400, // width
                100  // height
            );

            // Ensure the shape floats freely over the sheet
            wordArt.Placement = PlacementType.FreeFloating;

            // Lock the shape to prevent selection, movement, editing, and resizing
            wordArt.IsLocked = true;

            // Save the workbook with the locked WordArt watermark
            workbook.Save("WatermarkedLocked.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
