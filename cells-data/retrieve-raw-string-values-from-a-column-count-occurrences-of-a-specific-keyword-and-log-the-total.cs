// Title: Count keyword occurrences in a specific column of an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that opens an .xlsx file, reads all string values from a given column, and returns the total number of cells containing a specified keyword. | Modify the sample to perform a case‑insensitive search for the keyword and output the row numbers where matches are found. | Extend the program to write the total count and the matching row indices to a log file while preserving existing error handling.
// Common Searches: aspacells c# count how many times a word appears in column A of an Excel file | c# read raw string values from a specific Excel column using Aspose.Cells | how to perform case‑insensitive keyword search in an Excel worksheet with Aspose.Cells | log total occurrences of a text pattern from an Excel column in a .NET console app | iterate through used rows in a worksheet and count matches with Aspose.Cells
// Tags: count keyword occurrences Aspose.Cells | read column string values C# | case-insensitive text search Aspose.Cells | log keyword count to console .NET | iterate rows used range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an .xlsx file, reads string values from a designated column, counts cells that contain a given keyword, and writes the total count to the console, with basic file‑existence checks and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Column index to process (e.g., column A = 0)
            int columnIndex = 0;

            // Keyword to search for
            string keyword = "TargetKeyword";

            int totalCount = 0;

            // Iterate through all used rows in the worksheet
            int maxRow = sheet.Cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                Cell cell = sheet.Cells[row, columnIndex];

                // Retrieve the cell's string value
                string cellText = cell.StringValue;

                // Increment counter if the cell contains the keyword
                if (!string.IsNullOrEmpty(cellText) && cellText.Contains(keyword))
                {
                    totalCount++;
                }
            }

            Console.WriteLine($"Total occurrences of \"{keyword}\": {totalCount}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
