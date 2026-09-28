// Title: Validate Excel cell values by comparing raw StringValue to a predefined numeric string list using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates over a set of cells, reads each cell's StringValue, and determines whether it matches any entry in an array of expected numeric strings. | Show how to combine the IsNumericValue property with a string‑array lookup to confirm that a cell contains a numeric string in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# how to check if a cell's text matches a list of numeric strings | retrieve raw string value from Excel cell and verify against expected values using Aspose.Cells | using IsNumericValue with StringValue to validate numeric data in Aspose.Cells workbook | compare Excel cell content to predefined numeric strings in .NET application | validate cell data types and values in Aspose.Cells for .NET
// Tags: compare cell StringValue with expected list Aspose.Cells | use IsNumericValue property Aspose.Cells | validate numeric string cells C# | retrieve raw string from Excel cell Aspose | check cell content against predefined numeric strings .NET

using System;
using Aspose.Cells;

// The example creates a workbook, fills cells A1‑A5 with various string values, defines an array of expected numeric strings, then loops through each cell, reads its raw StringValue, checks if it exists in the expected array, evaluates the IsNumericValue flag, prints the results, and saves the file as ValidateCellContent.xlsx.
class ValidateCellContent
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate cells with various string values
        cells["A1"].PutValue("123");          // numeric string
        cells["A2"].PutValue("45.67");        // numeric string with decimal
        cells["A3"].PutValue("ABC");          // non‑numeric string
        cells["A4"].PutValue("2021-06-20");   // date string (treated as numeric after conversion)
        cells["A5"].PutValue("12/34/5678");   // invalid date string

        // Define the set of expected numeric strings
        string[] expectedNumericStrings = { "123", "45.67", "2021-06-20" };

        // Validate each cell by comparing its raw StringValue with the expected set
        foreach (string address in new[] { "A1", "A2", "A3", "A4", "A5" })
        {
            Cell cell = cells[address];
            string rawString = cell.StringValue;                     // retrieve raw string value
            bool matchesExpected = Array.Exists(expectedNumericStrings, s => s == rawString);
            bool isNumeric = cell.IsNumericValue;                    // additional numeric check

            Console.WriteLine($"{address}: Raw=\"{rawString}\" | MatchesExpected={matchesExpected} | IsNumericValue={isNumeric}");
        }

        // Save the workbook (saving rule)
        workbook.Save("ValidateCellContent.xlsx");
    }
}
