// Title: Set a custom print area that skips hidden rows and columns using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that identifies the first and last visible rows and columns in a worksheet and assigns the resulting range to PageSetup.PrintArea. | Create a reusable method that receives a Worksheet object, computes the visible cell range excluding hidden rows and columns, and returns the address string for setting the print area. | Show how to load an existing Excel file, apply a print area limited to visible cells, and save the workbook with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# set print area to visible cells only | How to exclude hidden rows and columns from Excel print area with Aspose.Cells .NET | Determine visible range in worksheet for custom print area using Aspose.Cells | Improve Excel printing performance by defining print area without hidden rows in C#
// Tags: Aspose.Cells set custom print area | skip hidden rows columns Aspose.Cells | PageSetup.PrintArea visible range .NET | determine first last visible row column C# | optimize Excel print performance Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, scans the first worksheet to find the first and last non‑hidden rows and columns, builds a range string from those indices, assigns it to the worksheet's PageSetup.PrintArea, and saves the modified file, thereby improving print performance by omitting hidden rows and columns.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Variables to hold the first and last visible row/column indices
            int firstVisibleRow = -1;
            int lastVisibleRow = -1;
            int firstVisibleColumn = -1;
            int lastVisibleColumn = -1;

            // Determine visible rows
            int maxRow = sheet.Cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                // Access the Row object via Cells.Rows collection
                if (!sheet.Cells.Rows[row].IsHidden)
                {
                    if (firstVisibleRow == -1) firstVisibleRow = row;
                    lastVisibleRow = row;
                }
            }

            // Determine visible columns
            int maxColumn = sheet.Cells.MaxDataColumn;
            for (int col = 0; col <= maxColumn; col++)
            {
                // Access the Column object via Cells.Columns collection
                if (!sheet.Cells.Columns[col].IsHidden)
                {
                    if (firstVisibleColumn == -1) firstVisibleColumn = col;
                    lastVisibleColumn = col;
                }
            }

            // If there is at least one visible row and column, set the custom print area
            if (firstVisibleRow != -1 && firstVisibleColumn != -1)
            {
                // Convert numeric indices to Excel cell references (e.g., 0 -> A)
                string startCell = CellsHelper.ColumnIndexToName(firstVisibleColumn) + (firstVisibleRow + 1);
                string endCell = CellsHelper.ColumnIndexToName(lastVisibleColumn) + (lastVisibleRow + 1);
                string printArea = $"{startCell}:{endCell}";

                // Apply the print area to the worksheet
                sheet.PageSetup.PrintArea = printArea;
            }

            // Save the workbook with the new print area
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
