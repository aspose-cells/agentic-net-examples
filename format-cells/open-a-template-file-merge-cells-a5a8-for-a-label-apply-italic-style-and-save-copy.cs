// Title: Merge cells A5:A8 in an Excel template, apply italic font style, and save as a copy using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells for C# to merge the range A5:A8 in an existing workbook, create an italic font style, apply it to the merged cells, and save the result as a new file. | Load a template workbook with Aspose.Cells, merge cells A5 through A8, set the font to italic via a custom style, and export the modified workbook as a separate copy.
// Common Searches: aspnet how to merge cells A5 to A8 and set italic font with Aspose.Cells | c# Aspose.Cells merge range and apply italic style then save copy | saving a modified Excel template as a new file using Aspose.Cells .NET | apply custom style to merged cells in Aspose.Cells C# example | merge cells and format font in Aspose.Cells workbook programmatically
// Tags: merge cells range Aspose.Cells C# | italic font formatting Aspose.Cells | save workbook copy Aspose.Cells | load Excel template Aspose.Cells | custom style creation Aspose.Cells

using Aspose.Cells;

// // This program loads "Template.xlsx", merges cells A5:A8 on the first worksheet, applies an italic font style to the merged region, and saves the result as "Template_Copy.xlsx" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the template workbook
        Workbook workbook = new Workbook("Template.xlsx");

        // Get the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells A5:A8 (rows 4-7, column 0)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns
        sheet.Cells.Merge(4, 0, 4, 1);

        // Create a style with italic font
        Style italicStyle = workbook.CreateStyle();
        italicStyle.Font.IsItalic = true;

        // Apply the italic style to the merged region (top‑left cell)
        sheet.Cells[4, 0].SetStyle(italicStyle);

        // Save the modified workbook as a copy
        workbook.Save("Template_Copy.xlsx");
    }
}
