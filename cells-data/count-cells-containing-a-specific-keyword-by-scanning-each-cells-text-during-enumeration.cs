// Title: Count Keyword Occurrences in Excel Cells with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, defines a target word, enumerates every instantiated cell in the first worksheet using Cells.GetEnumerator(), checks each Cell.StringValue for the word, tallies matches, and prints the total count.
// Keywords: Aspose.Cells | C# | .NET | keyword search in Excel | count cells containing text | enumerate worksheet cells | Cell.StringValue | Excel data analysis | text occurrence counter | programmatic Excel search
// Common Searches: Aspose.Cells count cells with specific text | How to find number of cells containing a word in Excel using C# | Enumerate all cells in a worksheet Aspose .NET | Search for keyword in Excel cells programmatically | Count occurrences of a string in an Excel sheet
// Developer Intent: Calculate how many cells in a worksheet include a particular keyword.
// Use Cases: Create a report showing the frequency of a term across a data sheet. | Validate that mandatory tags appear a required number of times before processing. | Trigger alerts when a keyword exceeds a predefined occurrence threshold.
// AI Prompts: Generate C# code with Aspose.Cells that counts cells containing "TargetKeyword" and displays the result. | Show how to modify the loop for a case‑insensitive search and ignore blank cells. | Explain how to extend the sample to aggregate keyword counts across every worksheet in the workbook.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsKeywordCount
{
    // Loads an Excel workbook, defines a target word, enumerates every instantiated cell in the first worksheet using Cells.GetEnumerator(), checks each Cell.StringValue for the word, tallies matches, and prints the total count.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Define the keyword to search for
            string keyword = "TargetKeyword";

            // Counter for cells containing the keyword
            int keywordCount = 0;

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Enumerate all instantiated cells in the worksheet
            IEnumerator cellEnumerator = worksheet.Cells.GetEnumerator();
            while (cellEnumerator.MoveNext())
            {
                // Cast the current object to Cell
                Cell cell = (Cell)cellEnumerator.Current;

                // Ensure the cell has a string representation
                string cellText = cell.StringValue;
                if (!string.IsNullOrEmpty(cellText) && cellText.Contains(keyword))
                {
                    keywordCount++;
                }
            }

            // Output the result
            Console.WriteLine($"Number of cells containing \"{keyword}\": {keywordCount}");

            // Optionally, save the workbook (no modifications made here)
            workbook.Save("output.xlsx");
        }
    }
}
