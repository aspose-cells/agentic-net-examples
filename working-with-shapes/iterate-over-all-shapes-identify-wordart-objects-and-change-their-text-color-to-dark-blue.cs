// Title: C# – Change WordArt Font Color to Dark Blue in All Worksheets Using Aspose.Cells
// Description: Loads a workbook, iterates every worksheet and its Shapes collection, detects WordArt objects via Shape.IsWordArt, sets their Font.Color to Color.DarkBlue, and saves the updated file.
// Keywords: Aspose.Cells C# WordArt color | Shape.IsWordArt Aspose.Cells | change WordArt font color .NET | iterate Excel shapes Aspose | bulk shape formatting Aspose.Cells | Color.DarkBlue Excel | C# Excel shape manipulation | Aspose.Cells API font color
// Common Searches: How to set WordArt text color in Excel with Aspose.Cells C# | Iterate all shapes in a workbook and change WordArt font color | Batch update WordArt colors across multiple worksheets using Aspose.Cells | Aspose.Cells example for changing WordArt to dark blue | C# code to detect WordArt and modify its font color in Excel
// Developer Intent: Locate every WordArt shape in an Excel workbook and apply a dark‑blue font color programmatically.
// Use Cases: Enforce corporate dark‑blue branding on all WordArt objects before publishing a multi‑sheet report. | Standardize visual appearance of generated dashboards by converting WordArt text to a single color. | Modernize legacy spreadsheets by automatically updating WordArt font colors to match new design guidelines.
// AI Prompts: Write C# code with Aspose.Cells that changes the font color of all WordArt shapes to a custom RGB value. | Explain how to identify WordArt shapes in a worksheet and modify their text formatting using Shape.IsWordArt. | Show a script that processes a folder of Excel files and batch‑updates WordArt text colors to dark blue with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, iterates every worksheet and its Shapes collection, detects WordArt objects via Shape.IsWordArt, sets their Font.Color to Color.DarkBlue, and saves the updated file.
class ChangeWordArtColor
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Identify WordArt shapes
                if (shape.IsWordArt)
                {
                    // Change the text color to dark blue
                    shape.Font.Color = Color.DarkBlue;
                }
            }
        }

        // Save the workbook with the updated WordArt colors
        workbook.Save("output.xlsx");
    }
}
