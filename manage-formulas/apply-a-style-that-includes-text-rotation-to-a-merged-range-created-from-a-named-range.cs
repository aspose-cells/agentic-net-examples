// Title: Rotate Text in a Merged Named Range with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, defines a UnionRange (A1:B2), names it "MyNamedRange", merges the cells, inserts "Rotated Text", builds a Style with a 45° RotationAngle, enables the rotation flag, applies the style to the merged range, and saves the file as MergedNamedRangeWithRotation.xlsx.
// Keywords: Aspose.Cells rotate text merged cells | C# text rotation named range | Aspose.Cells UnionRange style | merged cell rotation Aspose .NET | apply StyleFlag rotation Aspose.Cells
// Common Searches: how to rotate text in a merged named range using Aspose.Cells C# | apply rotation angle to UnionRange Aspose.Cells | merge cells, assign name, and set text angle Aspose.Cells .NET | Aspose.Cells style flag rotation for merged cells
// Developer Intent: Add a 45‑degree text rotation to a merged range that has been assigned a name.
// Use Cases: Design a spreadsheet header that spans several columns, merge it, name it, and rotate the text for a compact layout. | Create category labels in dashboards where merged cells are rotated to conserve horizontal space. | Programmatically generate reports with rotated text in merged named ranges for improved visual hierarchy.
// AI Prompts: Show C# code that merges a named range and rotates its text 45 degrees using Aspose.Cells. | Explain how to use Style and StyleFlag to set a rotation angle on a UnionRange in Aspose.Cells for .NET. | Provide an example of applying multiple style attributes (font, color, rotation) to a merged named range with Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, defines a UnionRange (A1:B2), names it "MyNamedRange", merges the cells, inserts "Rotated Text", builds a Style with a 45° RotationAngle, enables the rotation flag, applies the style to the merged range, and saves the file as MergedNamedRangeWithRotation.xlsx.
class ApplyRotationToMergedNamedRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cells that will form the named range (A1:B2)
            AsposeRange baseRange = worksheet.Cells.CreateRange("A1:B2");

            // Create a UnionRange from the base range
            UnionRange unionRange = worksheet.Cells
                .CreateRange("A1:B2")
                .UnionRanges(new AsposeRange[] { baseRange });

            // Assign a name to the union range
            unionRange.Name = "MyNamedRange";

            // Merge the cells in the union range
            unionRange.Merge();

            // Put a sample value into the merged cell (top‑left cell of the range)
            // The overload requires flags for conversion and formula handling
            unionRange.PutValue("Rotated Text", false, false);

            // Create a style and set the text rotation angle
            Style style = workbook.CreateStyle();
            style.RotationAngle = 45; // rotate text 45 degrees

            // Enable the rotation flag so the rotation is applied
            StyleFlag flag = new StyleFlag();
            flag.Rotation = true;

            // Apply the style with rotation to the merged range
            unionRange.ApplyStyle(style, flag);

            // Save the workbook
            string outputPath = "MergedNamedRangeWithRotation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
