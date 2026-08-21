// Title: C# – Create a Composite Named Range with UnionRange and Apply a Custom Style using Aspose.Cells
// Description: This example shows how to initialize a Workbook, define three separate ranges (A1:B2, C3:D4, E5:F6), merge them into a UnionRange, assign the name "MyCompositeRange", create a custom style (light‑green fill, bold dark‑blue centered text), apply the style to the composite range, and save the file as CompositeNamedRange.xlsx.
// Keywords: Aspose.Cells UnionRange C# | .NET composite named range | apply custom style Aspose.Cells | union multiple cell ranges | named range with style | Aspose.Cells SetStyle example
// Common Searches: how to union non‑contiguous ranges in Aspose.Cells | create a named range from multiple areas .NET | apply formatting to a UnionRange in Aspose.Cells | save workbook after styling composite range
// Developer Intent: Combine several non‑adjacent cell blocks into a single named range and format that range with a custom style in a .NET application.
// Use Cases: Define a reusable range that spans header and data sections for complex formulas. | Apply uniform formatting to disjoint cells in financial or reporting worksheets. | Provide a multi‑area data source for charts or pivot tables without manual range selection.
// AI Prompts: Generate C# code that unions A1:B2, C3:D4, and E5:F6 with Aspose.Cells, names the composite range, applies a bold centered style with a light‑green background, and saves the workbook. | Explain the UnionRange class in Aspose.Cells, covering creation, naming, and style application. | Suggest best practices for error handling when creating and styling a composite named range in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// This example shows how to initialize a Workbook, define three separate ranges (A1:B2, C3:D4, E5:F6), merge them into a UnionRange, assign the name "MyCompositeRange", create a custom style (light‑green fill, bold dark‑blue centered text), apply the style to the composite range, and save the file as CompositeNamedRange.xlsx.
class CreateCompositeNamedRange
{
    static void Main()
    {
        try
        {
            // Initialize a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Optional: add some sample data to visualize the ranges
            worksheet.Cells["A1"].PutValue("R1");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["C3"].PutValue("R2");
            worksheet.Cells["D4"].PutValue(20);
            worksheet.Cells["E5"].PutValue("R3");
            worksheet.Cells["F6"].PutValue(30);

            // Define three separate ranges
            AsposeRange range1 = worksheet.Cells.CreateRange("A1:B2");
            AsposeRange range2 = worksheet.Cells.CreateRange("C3:D4");
            AsposeRange range3 = worksheet.Cells.CreateRange("E5:F6");

            // Union the three ranges into a composite UnionRange
            UnionRange compositeRange = range1.UnionRanges(new AsposeRange[] { range2, range3 });

            // Assign a custom name to the composite range
            compositeRange.Name = "MyCompositeRange";

            // Create a custom style
            Style customStyle = workbook.CreateStyle();
            customStyle.ForegroundColor = Color.LightGreen;
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.Font.IsBold = true;
            customStyle.Font.Color = Color.DarkBlue;
            customStyle.HorizontalAlignment = TextAlignmentType.Center;
            customStyle.VerticalAlignment = TextAlignmentType.Center;

            // Apply the custom style to the composite range
            compositeRange.SetStyle(customStyle);

            // Save the workbook
            workbook.Save("CompositeNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
