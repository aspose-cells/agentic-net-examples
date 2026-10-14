// Title: Count cells with non‑empty HTML markup in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, scans each worksheet's used range, and returns the number of cells whose string value contains HTML tags. | Create a helper method that checks whether a string includes HTML markup by looking for opening and closing angle brackets, and use it while counting cells. | Improve performance by limiting the loops to MaxDataRow and MaxDataColumn for each worksheet when counting cells that contain HTML.
// Common Searches: aspocells count cells that contain html tags in c# | how to detect html markup in excel cell values using aspocells .net | iterate only used range of worksheets with aspocells to improve performance | c# sample to count non‑empty html cells in an xlsx workbook | aspocells example for scanning cells for html content
// Tags: html cell count aspocells | html markup detection in excel cells aspocells | used range iteration aspocells | cell string value analysis aspocells | maxdatarow maxdatacolumn loop aspocells | excel workbook cell scanning aspocells

using Aspose.Cells;
using System;

// The program loads an Excel workbook with Aspose.Cells, iterates through each worksheet's used range, checks each cell for a non‑empty string, uses a simple tag‑search to identify HTML markup, increments a counter for matching cells, and prints the total number of cells containing HTML content.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        int htmlCellCount = 0;

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range to limit iteration
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell contains a non‑empty string
                    if (cell.Value != null && cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue.Trim();

                        // Simple detection of HTML content (looks for <...> tags)
                        if (!string.IsNullOrEmpty(text) && ContainsHtml(text))
                        {
                            htmlCellCount++;
                        }
                    }
                }
            }
        }

        Console.WriteLine($"Number of cells with non‑empty HTML content: {htmlCellCount}");
    }

    // Helper method to detect presence of HTML tags in a string
    static bool ContainsHtml(string text)
    {
        int start = text.IndexOf('<');
        int end = text.IndexOf('>', start + 1);
        return start >= 0 && end > start;
    }
}
