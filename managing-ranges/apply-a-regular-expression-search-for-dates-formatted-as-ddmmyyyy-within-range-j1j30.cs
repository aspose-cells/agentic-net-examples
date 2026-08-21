// Title: C# Aspose.Cells: Regex search for dd/MM/yyyy dates in range J1:J30
// Description: Creates a workbook, fills column J with sample data, configures FindOptions for regular‑expression matching, restricts the search to J1:J30, and iterates with Cells.Find to list every cell whose value matches the pattern \d{2}/\d{2}/\d{4}. The workbook can be saved after the scan.
// Keywords: Aspose.Cells regex search | C# find dates in Excel | FindOptions SetRange | dd/MM/yyyy pattern | search specific column J | .NET Excel date validation
// Common Searches: Aspose.Cells find cells with date format dd/MM/yyyy | C# regex search in Excel range J1:J30 | How to limit Aspose.Cells Find to a column | Use FindOptions to locate date strings in Excel | Aspose.Cells regular expression example C#
// Developer Intent: Identify every cell that contains a date formatted as dd/MM/yyyy within the J1:J30 area.
// Use Cases: Validate that a column contains only correctly formatted dates before data import. | Extract rows with valid date strings for date‑driven reporting. | Flag or highlight cells with malformed dates during data‑cleansing.
// AI Prompts: Generate C# code using Aspose.Cells to highlight cells that match a dd/MM/yyyy regex in a given range. | Show how to extend FindOptions to search multiple columns for the same date pattern. | Explain how to modify the regex to accept single‑digit day or month values (e.g., d/M/yyyy).

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills column J with sample data, configures FindOptions for regular‑expression matching, restricts the search to J1:J30, and iterates with Cells.Find to list every cell whose value matches the pattern \d{2}/\d{2}/\d{4}. The workbook can be saved after the scan.
    public class RegexDateSearchInRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data in column J (index 9)
                cells["J1"].PutValue("Date");
                cells["J2"].PutValue("12/05/2023");
                cells["J3"].PutValue("InvalidDate");
                cells["J4"].PutValue("01/01/2022");
                cells["J5"].PutValue("23/12/2021");
                // Additional rows can be filled as needed...

                // Define find options for regex search
                FindOptions options = new FindOptions
                {
                    RegexKey = true,                         // Enable regex parsing
                    LookAtType = LookAtType.EntireContent,   // Exact match of the cell content
                    LookInType = LookInType.Values           // Search in cell values
                };

                // Restrict the search to the range J1:J30
                CellArea searchArea = new CellArea
                {
                    StartRow = 0,      // Row 1 (zero‑based)
                    EndRow = 29,       // Row 30
                    StartColumn = 9,   // Column J (zero‑based)
                    EndColumn = 9
                };
                options.SetRange(searchArea);

                // Regular expression for dates in dd/MM/yyyy format
                string dateRegex = @"\d{2}/\d{2}/\d{4}";

                // Find all matching cells
                Cell previousCell = null;
                Cell foundCell = cells.Find(dateRegex, previousCell, options);
                while (foundCell != null)
                {
                    Console.WriteLine($"Found date at {foundCell.Name}: {foundCell.StringValue}");
                    previousCell = foundCell;
                    foundCell = cells.Find(dateRegex, previousCell, options);
                }

                // Save the workbook (optional)
                workbook.Save("RegexDateSearchResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                RegexDateSearchInRange.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
