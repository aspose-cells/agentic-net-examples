// Title: How to fill a merged cell range with the workbook's Accent3 theme color using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set solid fill of merged cells to workbook theme Accent3 | apply workbook theme color to a merged range in Aspose.Cells .NET | C# merge cells and use theme Accent3 as background fill with Aspose.Cells | retrieve theme Accent3 color and style merged cells in Aspose.Cells | Aspose.Cells example for applying solid theme color to merged cells
// Tags: merged range solid fill Aspose.Cells | workbook theme accent3 color C# | styleflag cell shading Aspose.Cells | create range apply style Aspose.Cells | excel theme color fill .NET

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, merges cells B2:D4, obtains the workbook's Accent3 theme color, sets it as a solid foreground fill on the merged range using a StyleFlag for cell shading, and saves the file as MergedAccent3.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to merge (B2:D4)
            int firstRow = 1;      // Excel row 2 (zero‑based index)
            int firstColumn = 1;   // Excel column B (zero‑based index)
            int totalRows = 3;     // Rows 2,3,4
            int totalColumns = 3;  // Columns B,C,D

            // Merge the cells across the defined rows and columns
            sheet.Cells.Merge(firstRow, firstColumn, totalRows, totalColumns);

            // Get the style of the top‑left cell of the merged range
            Style style = sheet.Cells[firstRow, firstColumn].GetStyle();

            // Retrieve the workbook's theme Accent3 color
            Color accent3 = workbook.GetThemeColor(ThemeColorType.Accent3);

            // Apply the Accent3 color to the fill of the merged cells
            style.ForegroundColor = accent3;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the entire merged range
            AsposeRange mergedRange = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);
            mergedRange.ApplyStyle(style, new StyleFlag { CellShading = true });

            // Save the workbook
            workbook.Save("MergedAccent3.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
