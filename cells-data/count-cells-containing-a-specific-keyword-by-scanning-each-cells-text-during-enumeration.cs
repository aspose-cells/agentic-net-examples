// Title: Count cells that contain a specific keyword in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells count cells with a given text in an XLSX file | how to enumerate all worksheets and find keyword occurrences with Aspose.Cells | scan used range for a word and get total matches using Aspose.Cells in .NET | Aspose.Cells retrieve number of cells that include a specific string C# | search text in Excel cells and count matches with Aspose.Cells API
// Tags: keyword occurrence counting Aspose.Cells | worksheet enumeration used range Aspose.Cells | string cell scanning Aspose.Cells C# | excel workbook loading error handling Aspose.Cells | text match within cells Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an XLSX workbook with Aspose.Cells, safely handling missing files and loading errors, iterates through every worksheet's used range, checks each string‑type cell for a specified keyword, increments a counter for each match, and outputs the total number of cells containing the keyword.
class Program
{
    static void Main()
    {
        // Path to the Excel file and the keyword to search for
        string filePath = "input.xlsx";
        string keyword = "TargetKeyword";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook (load rule)
            workbook = new Workbook(filePath);
        }
        catch (CellsException ex)
        {
            // Handle cases such as password‑protected files or corrupted workbooks
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            // General exception handling
            Console.WriteLine($"Unexpected error: {ex.Message}");
            return;
        }

        // Counter for cells containing the keyword
        int count = 0;

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Scan each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue;

                        // Increment count if the cell text contains the keyword
                        if (!string.IsNullOrEmpty(text) && text.Contains(keyword))
                        {
                            count++;
                        }
                    }
                }
            }
        }

        // Output the total count
        Console.WriteLine($"Number of cells containing \"{keyword}\": {count}");
    }
}
