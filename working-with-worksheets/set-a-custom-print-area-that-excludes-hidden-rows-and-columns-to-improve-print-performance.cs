// Title: Set a print area that excludes hidden rows and columns using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that scans a worksheet for the first and last visible rows and columns and assigns that range to PageSetup.PrintArea. | Write a program that loads an existing Excel file, removes hidden rows/columns from the printable area, and saves the workbook with the updated print settings. | Create a method that calculates the visible data bounds based on MaxDataRow/MaxDataColumn and configures the worksheet's print area accordingly.
// Common Searches: how to define a print area in Aspose.Cells that ignores hidden rows and columns | Aspose.Cells .NET set printable range based on visible cells only | C# exclude hidden rows from Excel print area using Aspose.Cells | optimize Excel printing speed by limiting print area to non‑hidden rows Aspose | PageSetup.PrintArea custom range for visible data Aspose.Cells
// Tags: Aspose.Cells set print area excluding hidden rows | visible rows print range Aspose.Cells .NET | PageSetup.PrintArea hidden columns handling | optimize Excel print performance with Aspose.Cells | calculate visible data bounds Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads a workbook, determines the first and last non‑hidden rows and columns by checking the IsHidden property, builds a cell range string, assigns it to the worksheet's PageSetup.PrintArea, and saves the updated file.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the range of rows that actually contain data
            int maxRow = sheet.Cells.MaxDataRow;
            int maxColumn = sheet.Cells.MaxDataColumn;

            // If the worksheet is empty, there is nothing to set as print area
            if (maxRow < 0 || maxColumn < 0)
            {
                Console.WriteLine("Worksheet contains no data.");
                workbook.Save(outputPath);
                return;
            }

            // Find first and last visible rows
            int firstVisibleRow = -1;
            int lastVisibleRow = -1;
            for (int i = 0; i <= maxRow; i++)
            {
                if (!sheet.Cells.Rows[i].IsHidden)
                {
                    if (firstVisibleRow == -1) firstVisibleRow = i;
                    lastVisibleRow = i;
                }
            }

            // Find first and last visible columns
            int firstVisibleColumn = -1;
            int lastVisibleColumn = -1;
            for (int j = 0; j <= maxColumn; j++)
            {
                if (!sheet.Cells.Columns[j].IsHidden)
                {
                    if (firstVisibleColumn == -1) firstVisibleColumn = j;
                    lastVisibleColumn = j;
                }
            }

            // Set custom print area if visible rows/columns were found
            if (firstVisibleRow != -1 && firstVisibleColumn != -1)
            {
                string startCell = CellsHelper.CellIndexToName(firstVisibleRow, firstVisibleColumn);
                string endCell = CellsHelper.CellIndexToName(lastVisibleRow, lastVisibleColumn);
                sheet.PageSetup.PrintArea = $"{startCell}:{endCell}";
            }

            // Save the workbook with the new print area
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
