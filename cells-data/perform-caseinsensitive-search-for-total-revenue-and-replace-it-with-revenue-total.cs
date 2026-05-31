using System;
using Aspose.Cells;

namespace AsposeCellsReplaceExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace "input.xlsx" with your source file)
            Workbook workbook = new Workbook("input.xlsx");

            // Configure replace options for case‑insensitive search
            ReplaceOptions options = new ReplaceOptions
            {
                CaseSensitive = false,               // ignore case
                MatchEntireCellContents = false      // replace even if the text is part of a larger string
            };

            // Perform the replacement: "total revenue" → "Revenue Total"
            int replacedCount = workbook.Replace("total revenue", "Revenue Total", options);

            Console.WriteLine($"Number of replacements made: {replacedCount}");

            // Save the modified workbook (replace "output.xlsx" with your desired output file)
            workbook.Save("output.xlsx");
        }
    }
}