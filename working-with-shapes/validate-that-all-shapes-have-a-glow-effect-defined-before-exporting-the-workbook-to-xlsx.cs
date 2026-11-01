// Title: C# – Validate that every shape has a glow effect before saving an Aspose.Cells workbook to XLSX
// AI Prompts: Loop through all worksheets in a Workbook, inspect each Shape's Glow property, and raise an InvalidOperationException if Glow is null or its Size equals zero. | Once every shape is confirmed to have a defined glow effect, invoke Workbook.Save with SaveFormat.Xlsx to export the file.
// Common Searches: how to ensure all shapes have a glow effect in Aspose.Cells before exporting to xlsx | c# Aspose.Cells validate shape glow property and throw exception | iterate over worksheet shapes and check glow size using Aspose.Cells .NET | save workbook as xlsx only after shape formatting validation in Aspose.Cells | exception handling for missing shape glow in Excel automation with Aspose.Cells
// Tags: shape glow validation Aspose.Cells | iterate worksheet shapes C# | check shape glow property .NET | save workbook as xlsx after validation | throw exception missing shape glow

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel file, iterates through each worksheet and its shapes, verifies that each shape's Glow property is non‑null and its Size is greater than zero, throws an InvalidOperationException if any shape lacks a glow effect, and finally saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each shape on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Validate that a glow effect is defined (non‑null and size > 0)
                if (shape.Glow == null || shape.Glow.Size == 0)
                {
                    // Throw an exception or handle the validation failure as needed
                    throw new InvalidOperationException(
                        $"Shape '{shape.Name}' on sheet '{sheet.Name}' does not have a glow effect defined.");
                }
            }
        }

        // Export the workbook to XLSX after successful validation
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
