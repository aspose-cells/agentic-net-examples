// Title: Perform a case‑insensitive find‑and‑replace of 'total revenue' with 'Revenue Total' in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells C# API to replace every occurrence of the phrase 'total revenue' with 'Revenue Total' across all worksheets, ignoring case. | Show how to configure ReplaceOptions for a case‑insensitive, partial‑match text substitution in an Excel file with Aspose.Cells. | Generate C# code that loads an .xlsx file, executes a case‑insensitive replace operation, and saves the modified workbook.
// Common Searches: asp.net replace text in Excel cells case insensitive Aspose.Cells | c# Aspose.Cells find and replace string ignoring case in workbook | how to change 'total revenue' to 'Revenue Total' in all sheets using Aspose.Cells | replace partial text in .xlsx with Aspose.Cells C# example
// Tags: Aspose.Cells ReplaceOptions ignore case | C# workbook text substitution | partial cell match replace .xlsx | Excel find replace across worksheets | string substitution using Aspose.Cells

using System;
using Aspose.Cells;

// The example creates (or loads) a workbook, writes sample data, sets up ReplaceOptions to ignore case and allow partial matches, replaces every occurrence of "total revenue" with "Revenue Total", outputs the number of replacements, and saves the updated file as Output.xlsx.
class ReplaceTotalRevenue
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // For loading: new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Example data – can be removed when using a real workbook
        sheet.Cells["A1"].PutValue("Total Revenue");
        sheet.Cells["A2"].PutValue("total revenue for Q1");
        sheet.Cells["A3"].PutValue("Net profit");

        // Configure replace options for case‑insensitive replacement
        ReplaceOptions options = new ReplaceOptions
        {
            CaseSensitive = false,               // ignore case
            MatchEntireCellContents = false      // allow partial matches within cells
        };

        // Perform the replacement
        int replacedCount = workbook.Replace("total revenue", "Revenue Total", options);
        Console.WriteLine($"Replacements made: {replacedCount}");

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
