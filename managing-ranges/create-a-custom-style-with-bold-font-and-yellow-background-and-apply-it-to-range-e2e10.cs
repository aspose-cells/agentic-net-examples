// Title: Apply a bold font and yellow background style to cells E2‑E10 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a style with bold text and a solid yellow fill in Aspose.Cells, then applies it to the range E2:E10. | Show how to configure a StyleFlag to apply only font boldness and cell shading when styling a specific range in an Aspose.Cells workbook. | Provide the steps to save the workbook after applying the custom style to the selected cells.
// Common Searches: asp.net aspose.cells set bold font and yellow background for range E2:E10 | c# aspose.cells apply custom style to specific cells | using StyleFlag to limit style attributes in Aspose.Cells | create solid fill style with bold text in Aspose.Cells .NET | how to save workbook after styling cells with Aspose.Cells
// Tags: apply custom style to cell range Aspose.Cells | bold font solid yellow fill Aspose.Cells | StyleFlag font bold cell shading .NET | create and apply style to E2:E10 Aspose.Cells | save workbook after styling Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program creates a new workbook, defines a custom style with bold text and a solid yellow background, uses a StyleFlag to apply only font boldness and cell shading, applies the style to cells E2 through E10 on the first worksheet, and saves the file as StyledRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a custom style with bold font and yellow background
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.IsBold = true;
            customStyle.ForegroundColor = Color.Yellow;
            customStyle.Pattern = BackgroundType.Solid;

            // Define which style attributes to apply
            StyleFlag flag = new StyleFlag
            {
                FontBold = true,
                CellShading = true
            };

            // Apply the style to the range E2:E10
            AsposeRange range = sheet.Cells.CreateRange("E2", "E10");
            range.ApplyStyle(customStyle, flag);

            // Save the workbook
            string outputPath = "StyledRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
