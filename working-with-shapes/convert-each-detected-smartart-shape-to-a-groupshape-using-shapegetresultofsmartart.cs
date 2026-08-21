// Title: Convert SmartArt to GroupShape in Excel using Aspose.Cells for .NET (C#)
// Description: Loads an Excel file, scans each worksheet for SmartArt shapes (IsSmartArt), converts them to GroupShape objects with Shape.GetResultOfSmartArt, optionally repositions them, and saves the workbook with OoxmlSaveOptions.UpdateSmartArt enabled.
// Keywords: Aspose.Cells SmartArt conversion | Shape.GetResultOfSmartArt C# | GroupShape Excel .NET | UpdateSmartArt OoxmlSaveOptions | iterate worksheet shapes Aspose | C# Excel shape manipulation | global .NET developers
// Common Searches: How to change SmartArt to GroupShape with Aspose.Cells | C# code for detecting SmartArt in Excel worksheets | Save Excel after modifying SmartArt using Aspose | Move converted GroupShape position Aspose.Cells
// Developer Intent: Replace every SmartArt object in an Excel workbook with an editable GroupShape and optionally adjust its coordinates.
// Use Cases: Enable full editing (resize, rotate, style) of formerly SmartArt graphics. | Prevent overlap by shifting the new GroupShape after conversion. | Ensure the workbook retains the changes by using UpdateSmartArt during save.
// AI Prompts: Generate a C# routine that iterates all worksheets, converts each SmartArt to a GroupShape with Shape.GetResultOfSmartArt, and logs null conversions. | Create code that records the original SmartArt name and the width/height of the resulting GroupShape. | Write a unit test that confirms a SmartArt shape becomes a GroupShape and that the file is saved with UpdateSmartArt set to true.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel file, scans each worksheet for SmartArt shapes (IsSmartArt), converts them to GroupShape objects with Shape.GetResultOfSmartArt, optionally repositions them, and saves the workbook with OoxmlSaveOptions.UpdateSmartArt enabled.
class ConvertSmartArtToGroupShape
{
    static void Main()
    {
        // Load the workbook that contains SmartArt shapes
        Workbook workbook = new Workbook("input.xlsx");

        // Process each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Examine each shape on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Check if the shape is a SmartArt object
                if (shape.IsSmartArt)
                {
                    // Convert the SmartArt shape to a GroupShape
                    GroupShape groupShape = shape.GetResultOfSmartArt();

                    // Optional: manipulate the resulting group shape
                    if (groupShape != null)
                    {
                        // Example: shift the group shape slightly
                        groupShape.Left += 10;
                        groupShape.Top += 10;
                    }
                }
            }
        }

        // Save the workbook, ensuring SmartArt updates are written back
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.UpdateSmartArt = true;
        workbook.Save("output.xlsx", saveOptions);
    }
}
