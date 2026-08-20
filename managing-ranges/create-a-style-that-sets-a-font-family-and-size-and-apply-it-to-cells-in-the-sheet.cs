// Title: Aspose.Cells .NET Example: Create a Font Style and Apply It to a Cell Range (C#)
// Description: This C# sample demonstrates how to create a custom Style with a specific font family and size using Aspose.Cells, configure a StyleFlag to limit the applied properties, and apply the style to a range of cells (A1:C3) before saving the workbook as StyledSheet.xlsx.
// Keywords: Aspose.Cells C# style example | set font family Aspose.Cells | apply font size to range Aspose.Cells | CreateStyle Aspose.Cells .NET | StyleFlag font properties | C# Excel styling Aspose | GitHub Aspose.Cells sample | Excel font style programmatically
// Common Searches: how to set font name and size for a range in Aspose.Cells .NET | apply custom style to multiple cells using Aspose.Cells C# | Aspose.Cells StyleFlag usage example | C# code to style Excel cells with Aspose | Aspose.Cells create and apply font style
// Developer Intent: Create a reusable font style (family and size) and apply it to a selected cell range in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design a header row with a larger, branded font across the top of a report. | Standardize body text formatting (e.g., Calibri 12pt) for all data cells in a financial statement. | Apply the same typography to multiple worksheets to ensure visual consistency across a workbook.
// AI Prompts: Generate C# code with Aspose.Cells that creates an Arial 12pt style and applies it to the range B2:D5. | Show how to extend the example to also set font color, bold, and underline while using StyleFlag. | Explain how to apply one custom style to several non‑contiguous ranges (e.g., A1:A5, C1:C5) in the same worksheet.

using System;
using Aspose.Cells;

// This C# sample demonstrates how to create a custom Style with a specific font family and size using Aspose.Cells, configure a StyleFlag to limit the applied properties, and apply the style to a range of cells (A1:C3) before saving the workbook as StyledSheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a style and set the desired font family and size
            Style style = workbook.CreateStyle();
            style.Font.Name = "Calibri";
            style.Font.Size = 14;

            // Define which style properties should be applied
            StyleFlag flag = new StyleFlag
            {
                FontName = true,
                FontSize = true
            };

            // Apply the style to a range of cells (A1:C3 in this example)
            Aspose.Cells.Range range = cells.CreateRange("A1", "C3");
            range.ApplyStyle(style, flag);

            // Add sample data to see the style in effect
            cells["A1"].PutValue("Sample");
            cells["B2"].PutValue("Text");
            cells["C3"].PutValue(123);

            // Save the workbook
            workbook.Save("StyledSheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
