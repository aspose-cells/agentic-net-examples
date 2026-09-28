// Title: How to set the default workbook font to Arial 10pt and apply it to all cells with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Workbook, sets its default style font to Arial 10pt, and applies this style to every cell in each worksheet, using a 100‑row by 26‑column fallback for empty sheets. | Write a method with Aspose.Cells that defines an Arial 10pt Font style, uses a StyleFlag to apply only the font to the full used range of all worksheets, and saves the workbook.
// Common Searches: Aspose.Cells C# set workbook default font to Arial 10 | apply a font style to all cells in an Excel file using Aspose.Cells .NET | how to change font for empty worksheets in Aspose.Cells | C# Aspose.Cells apply style to entire sheet including unused rows | set global font for new Excel workbook with Aspose.Cells
// Tags: default workbook font Aspose.Cells .NET | apply style to entire worksheet Aspose.Cells | StyleFlag font only Aspose.Cells | global Arial 10pt font Excel Aspose | handle empty sheet range Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, changes the default style to Arial 10pt, iterates through each worksheet, determines the used range (or defaults to rows 0‑99 and columns 0‑25 for empty sheets), creates a matching style, applies it with a Font‑only StyleFlag, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook's default font to Arial, size 10
        Style defaultStyle = workbook.DefaultStyle;
        defaultStyle.Font.Name = "Arial";
        defaultStyle.Font.Size = 10;
        workbook.DefaultStyle = defaultStyle;

        // Apply the default font to all cells in each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the range to apply the style (used range or a default area if empty)
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;
            if (maxRow < 0 || maxCol < 0)
            {
                // Sheet is empty; define a reasonable default area
                maxRow = 99;   // rows 0-99 (100 rows)
                maxCol = 25;   // columns 0-25 (A-Z)
            }

            // Create a style with the desired font
            Style style = workbook.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 10;

            // Apply the style to the entire range
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = maxRow,
                EndColumn = maxCol
            };
            sheet.Cells.ApplyStyle(style, new StyleFlag { Font = true });
        }

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
