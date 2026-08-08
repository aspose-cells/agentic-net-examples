// Title: C# – Count case‑insensitive keyword occurrences in an Excel column with Aspose.Cells
// Description: Creates a workbook, writes sample strings to column A, defines a search term, scans each populated cell using its raw string value, increments a counter when the term is found (ignoring case), prints the total and saves the file as KeywordCount.xlsx.
// Keywords: Aspose.Cells C# keyword count | case insensitive text search Excel .NET | retrieve raw cell value Aspose.Cells | count word occurrences column | Excel data analysis Aspose.Cells | C# iterate worksheet cells | Aspose.Cells performance scanning | search string in Excel column
// Common Searches: Aspose.Cells count word in column C# | case‑insensitive search Excel worksheet using .NET | how to get raw cell text with Aspose.Cells | C# count occurrences of a string in Excel column | Aspose.Cells iterate rows to find keyword
// Developer Intent: Determine how many times a specific word or phrase appears in a chosen column of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Generate a term‑frequency report for a data column. | Validate that required keywords meet a minimum count in imported spreadsheets. | Build a summary sheet that aggregates keyword totals across multiple columns or worksheets.
// AI Prompts: Create a reusable C# method that accepts a Worksheet, column index, and search term, then returns the case‑insensitive occurrence count using Aspose.Cells. | Show how to loop through all worksheets in a workbook, count a given keyword in each column, and write the results to a new summary sheet. | Explain techniques to improve performance when scanning large Excel files for text matches with Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, writes sample strings to column A, defines a search term, scans each populated cell using its raw string value, increments a counter when the term is found (ignoring case), prints the total and saves the file as KeywordCount.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate column A with sample string values
        string[] sampleData = { "apple", "banana", "apple pie", "orange", "Apple", "grape", "pineapple" };
        for (int i = 0; i < sampleData.Length; i++)
        {
            sheet.Cells[i, 0].PutValue(sampleData[i]); // Column index 0 = column A
        }

        // Define the keyword to search for (case‑insensitive)
        string keyword = "apple";

        // Count occurrences of the keyword in column A
        int occurrenceCount = 0;
        int lastRow = sheet.Cells.MaxDataRow; // Last row that contains data
        for (int row = 0; row <= lastRow; row++)
        {
            string cellText = sheet.Cells[row, 0].StringValue;
            if (!string.IsNullOrEmpty(cellText) &&
                cellText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                occurrenceCount++;
            }
        }

        // Log the total count
        Console.WriteLine($"Total occurrences of \"{keyword}\" in column A: {occurrenceCount}");

        // Save the workbook
        workbook.Save("KeywordCount.xlsx");
    }
}
