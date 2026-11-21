// Title: Set a custom print area that excludes hidden rows and columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set print area only for visible rows and columns | exclude hidden rows from Excel print area programmatically using Aspose.Cells | determine visible range and assign PageSetup.PrintArea in .NET | C# Aspose.Cells print area without hidden rows and columns | how to create dynamic print area based on visible cells in Aspose.Cells
// Tags: Aspose.Cells set PrintArea for visible cells | skip hidden rows and columns in Excel print area | find first and last visible row and column C# | PageSetup.PrintArea dynamic range Aspose.Cells | custom print area based on visible range .NET

using System;
using Aspose.Cells;

// The example loads an Excel workbook, scans the first worksheet to locate the first and last visible rows and columns, builds the corresponding address range, assigns it to the worksheet's PageSetup.PrintArea, and saves the updated file.
class SetPrintAreaExcludingHidden
{
    static void Main()
    {
        // Load the workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (adjust as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the range of visible rows
        int firstVisibleRow = -1;
        int lastVisibleRow = -1;
        int maxRow = sheet.Cells.MaxDataRow; // last row that contains data
        for (int i = 0; i <= maxRow; i++)
        {
            if (!sheet.Cells.Rows[i].IsHidden)
            {
                if (firstVisibleRow == -1) firstVisibleRow = i;
                lastVisibleRow = i;
            }
        }

        // Determine the range of visible columns
        int firstVisibleCol = -1;
        int lastVisibleCol = -1;
        int maxCol = sheet.Cells.MaxDataColumn; // last column that contains data
        for (int j = 0; j <= maxCol; j++)
        {
            if (!sheet.Cells.Columns[j].IsHidden)
            {
                if (firstVisibleCol == -1) firstVisibleCol = j;
                lastVisibleCol = j;
            }
        }

        // If there are visible rows and columns, set the print area
        if (firstVisibleRow != -1 && firstVisibleCol != -1)
        {
            // Convert column indices to column letters (A, B, ...)
            string startColName = CellsHelper.ColumnIndexToName(firstVisibleCol);
            string endColName = CellsHelper.ColumnIndexToName(lastVisibleCol);

            // Row numbers are 1‑based in Excel addresses
            string startCell = $"{startColName}{firstVisibleRow + 1}";
            string endCell = $"{endColName}{lastVisibleRow + 1}";

            // Build the address string for the print area
            string printArea = $"{startCell}:{endCell}";

            // Apply the custom print area
            sheet.PageSetup.PrintArea = printArea;
        }
        else
        {
            // No visible cells found – optionally clear any existing print area
            sheet.PageSetup.PrintArea = string.Empty;
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
