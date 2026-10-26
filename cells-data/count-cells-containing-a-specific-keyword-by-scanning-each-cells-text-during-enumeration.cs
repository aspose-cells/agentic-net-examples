// Title: Count cells that contain a specific keyword in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells count cells that contain a given word in an Excel file | How to enumerate all worksheets and find keyword occurrences with Aspose.Cells | Case‑insensitive search for a string across all cells in a .xlsx using Aspose.Cells | Get total number of cells matching a text pattern in a workbook with Aspose.Cells .NET | Scan used range of Excel sheets for a specific keyword using Aspose.Cells API
// Tags: keyword search in Excel cells Aspose.Cells | count matching cells .NET Aspose.Cells | enumerate used range cells Aspose.Cells | case‑insensitive text match Excel C# | scan worksheet cells for specific string Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, iterates through each worksheet's used range, checks each non‑empty cell's displayed text for a case‑insensitive occurrence of a given keyword, increments a counter for each match, and outputs the total number of cells containing the keyword while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be processed
        string filePath = "input.xlsx";

        // Keyword to search for within cell texts
        string keyword = "target";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            int matchingCellCount = 0;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the cells collection of the current worksheet
                Cells cells = sheet.Cells;

                // Determine the used range to limit enumeration
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];

                        // Skip cells with no value
                        if (cell.Type == CellValueType.IsNull)
                            continue;

                        // Retrieve the cell's displayed text
                        string cellText = cell.StringValue;

                        // Check if the cell text contains the keyword (case‑insensitive)
                        if (!string.IsNullOrEmpty(cellText) &&
                            cellText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            matchingCellCount++;
                        }
                    }
                }
            }

            // Output the total count of cells containing the keyword
            Console.WriteLine($"Number of cells containing \"{keyword}\": {matchingCellCount}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors without crashing the application
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
