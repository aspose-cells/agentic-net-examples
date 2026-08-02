// Title: Apply Background Color to the Intersection of Two Named Ranges and Save as XLSX – Aspose.Cells for .NET
// Description: Creates a workbook, defines two overlapping named ranges, uses UnionRange.Intersect to locate the common cells, applies a solid yellow fill style to that intersection, and saves the result as an XLSX file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells UnionRange.Intersect | highlight overlapping ranges C# | apply fill color to intersected cells | save workbook as XLSX Aspose | named range intersection styling
// Common Searches: how to highlight intersecting named ranges Aspose.Cells | UnionRange.Intersect example C# | apply background color to overlapping cells .NET | save styled workbook as xlsx using Aspose
// Developer Intent: Format the cells that belong to both named ranges with a background color and export the workbook to XLSX.
// Use Cases: Mark overlapping data sections in automated reports. | Create visual cues for merged datasets in dashboards. | Emphasize intersecting zones when consolidating multiple spreadsheets.
// AI Prompts: Add a thick border around the intersected range in this Aspose.Cells sample. | Rewrite the code to read range names from a JSON file and use conditional formatting instead of a solid fill. | Explain how UnionRange.Intersect works and show best practices for null‑check handling.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, defines two overlapping named ranges, uses UnionRange.Intersect to locate the common cells, applies a solid yellow fill style to that intersection, and saves the result as an XLSX file with Aspose.Cells for .NET.
class ApplyIntersectionBackground
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Data1");
            sheet.Cells["B1"].PutValue("Data1");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data1");

            sheet.Cells["B1"].PutValue("Data2");
            sheet.Cells["C1"].PutValue("Data2");
            sheet.Cells["B2"].PutValue("Data2");
            sheet.Cells["C2"].PutValue("Data2");

            // Create the first range and assign a name
            AsposeRange range1 = sheet.Cells.CreateRange("A1:B2");
            range1.Name = "FirstRange";

            // Create the second range and assign a name
            AsposeRange range2 = sheet.Cells.CreateRange("B1:C2");
            range2.Name = "SecondRange";

            // Convert each Range to a UnionRange (required for UnionRange.Intersect)
            UnionRange union1 = sheet.Cells.CreateRange(range1.RefersTo)
                                         .UnionRanges(new AsposeRange[] { range1 });

            UnionRange union2 = sheet.Cells.CreateRange(range2.RefersTo)
                                         .UnionRanges(new AsposeRange[] { range2 });

            // Get the intersected area of the two UnionRanges
            UnionRange intersected = union1.Intersect(union2);

            if (intersected != null)
            {
                // Create a style with a solid yellow background
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = Color.Yellow;
                highlightStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the intersected area
                intersected.SetStyle(highlightStyle);
            }

            // Save the workbook as XLSX
            workbook.Save("IntersectedBackground.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
