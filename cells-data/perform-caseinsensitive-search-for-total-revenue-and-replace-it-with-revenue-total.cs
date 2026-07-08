using System;
using Aspose.Cells;

namespace AsposeCellsReplaceExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data containing different case variations of the target phrase
            sheet.Cells["A1"].PutValue("Total Revenue");
            sheet.Cells["A2"].PutValue("total revenue");
            sheet.Cells["A3"].PutValue("TOTAL REVENUE");
            sheet.Cells["A4"].PutValue("Net Total Revenue");

            // Configure replace options for case‑insensitive search
            ReplaceOptions options = new ReplaceOptions
            {
                CaseSensitive = false,          // ignore case
                MatchEntireCellContents = false // allow partial matches within a cell
            };

            // Perform the replacement
            int replacedCount = workbook.Replace("total revenue", "Revenue Total", options);
            Console.WriteLine($"Replacements made: {replacedCount}");

            // Save the workbook
            workbook.Save("ReplacedOutput.xlsx");
        }
    }
}