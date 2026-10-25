// Title: Set worksheet FitToPagesWide dynamically from column count with a configurable factor using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set FitToPagesWide based on number of used columns | calculate pages wide for Excel print layout using a factor in Aspose.Cells | dynamic print scaling in .NET workbook by dividing column count by custom factor | how to automatically adjust FitToPagesWide in Aspose.Cells according to column count | C# Aspose.Cells fit worksheet to pages wide using column count and factor
// Tags: fit-to-pages-wide calculation Aspose.Cells | worksheet print scaling by column count | configurable factor for Excel page layout | dynamic page fit for .xlsx using C# | Aspose.Cells workbook print setup automation

using Aspose.Cells;
using System;

// The example loads an existing workbook, determines the count of used columns in the first worksheet, computes the required pages wide by dividing that count by a configurable factor and rounding up, sets the worksheet's FitToPagesWide to this value while keeping FitToPagesTall at zero for automatic height adjustment, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Configurable factor that determines how many columns fit on one page
        int factor = 5; // Adjust this value as needed

        // Load an existing workbook (replace with actual path)
        Workbook workbook = new Workbook("input.xlsx");

        // Work with the first worksheet (modify as required)
        Worksheet sheet = workbook.Worksheets[0];

        // Get the index of the last column that contains data (0‑based)
        int lastColumnIndex = sheet.Cells.MaxDataColumn;

        // Calculate total number of used columns (1‑based count)
        int columnCount = lastColumnIndex + 1;

        // Determine how many pages wide the sheet should be.
        // Divide column count by the factor and round up to the next whole page.
        int pagesWide = (int)Math.Ceiling((double)columnCount / factor);
        if (pagesWide < 1) pagesWide = 1; // Ensure at least one page

        // Apply the FitToPagesWide setting; set FitToPagesTall to 0 to let height adjust automatically
        sheet.PageSetup.FitToPagesWide = pagesWide;
        sheet.PageSetup.FitToPagesTall = 0;

        // Save the modified workbook (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}
