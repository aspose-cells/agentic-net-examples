// Title: Merge cells L5:M7, apply italic gray fill style, and save as XLSX using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, merge the range L5:M7, define a style with italic font and a gray solid fill, apply the style to the merged range, and save the file as MergedStyled.xlsx using Aspose.Cells in C#. | Using Aspose.Cells for .NET, generate a style that sets Font.IsItalic = true and a gray background, apply this style to the merged cell block L5:M7, then export the workbook to an XLSX file.
// Common Searches: Aspose.Cells C# merge a rectangular range and set italic font with gray background | How to apply custom style to merged cells L5:M7 in a .NET Excel workbook | Saving a workbook with styled merged cells using Aspose.Cells for .NET | Create and apply italic gray style to a merged range in Aspose.Cells C# example
// Tags: merge cell range L5:M7 Aspose.Cells | italic font style Aspose.Cells C# | gray solid fill background Aspose.Cells | apply style to merged cells Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program creates a new workbook, merges cells L5:M7, defines a custom style with italic font and a gray solid fill, applies this style to the merged range, and saves the result as MergedStyled.xlsx.
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

            // Merge cells L5:M7 (zero‑based indices: row 4, column 11, 3 rows, 2 columns)
            sheet.Cells.Merge(4, 11, 3, 2);

            // Create a style with italic font and gray fill
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.IsItalic = true;
            customStyle.ForegroundColor = Color.Gray;
            customStyle.Pattern = BackgroundType.Solid;

            // Define which style elements to apply
            StyleFlag flag = new StyleFlag
            {
                FontItalic = true,
                CellShading = true
            };

            // Apply the style to the merged range L5:M7
            AsposeRange mergedRange = sheet.Cells.CreateRange("L5:M7");
            mergedRange.ApplyStyle(customStyle, flag);

            // Save the workbook
            workbook.Save("MergedStyled.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
