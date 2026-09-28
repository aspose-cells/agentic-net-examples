// Title: Count case‑insensitive occurrences of a keyword in a specific Excel column using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans column B of an Excel workbook and returns the number of cells that contain the word "Invoice" regardless of case. | Modify the example to also output the row indices where the keyword is found, using a List<int>. | Create a reusable method `int CountKeyword(string filePath, string keyword, int columnIndex)` that loads a workbook with Aspose.Cells and returns the occurrence count for any column.
// Common Searches: aspocells count how many times a word appears in column A c# | case insensitive text search in Excel column using Aspose.Cells .NET | retrieve string values from a single column and count keyword occurrences with Aspose.Cells | C# Aspose.Cells example to count keyword occurrences in a worksheet column | how to get total matches of a specific word in Excel column using Aspose.Cells
// Tags: keyword count column Aspose.Cells | case-insensitive search column Aspose.Cells | iterate column cells Aspose.Cells | load workbook process column Aspose.Cells | log occurrence count console C#

using System;
using Aspose.Cells;

// Loads an Excel workbook, iterates through a specified column, counts case‑insensitive occurrences of a given keyword in string cells, writes the total to the console, and optionally saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Keyword to search for
        string keyword = "Aspose";

        // Counter for occurrences
        int occurrenceCount = 0;

        // Determine the last row that contains data in the worksheet
        int lastRow = cells.MaxDataRow;

        // Iterate through each cell in column A (index 0)
        for (int row = 0; row <= lastRow; row++)
        {
            Cell cell = cells[row, 0]; // Column A

            // Process only string cells
            if (cell.Type == CellValueType.IsString)
            {
                string cellText = cell.StringValue;

                // Check if the cell text contains the keyword (case‑insensitive)
                if (!string.IsNullOrEmpty(cellText) &&
                    cellText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    occurrenceCount++;
                }
            }
        }

        // Log the total count
        Console.WriteLine($"Total occurrences of \"{keyword}\" in column A: {occurrenceCount}");

        // Save the workbook (optional, can be omitted if no changes are needed)
        workbook.Save("output.xlsx");
    }
}
