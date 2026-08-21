// Title: C# utility to list worksheets with QueryTables using Aspose.Cells
// Description: A command‑line program that loads an Excel workbook with Aspose.Cells, scans each worksheet, and prints the names of those that contain one or more QueryTables.
// Keywords: Aspose.Cells QueryTables list worksheets | C# enumerate QueryTables | find sheets with data connections Aspose | list worksheet names containing QueryTables | Aspose.Cells workbook query tables
// Common Searches: how to list worksheets that have QueryTables in Aspose.Cells | C# code to get sheet names with data connections | enumerate QueryTables per worksheet Aspose.Cells .NET | list Excel sheets containing query tables programmatically
// Developer Intent: Identify and output the names of all worksheets that contain at least one QueryTable in a specified workbook.
// Use Cases: Audit workbooks to ensure only approved sheets contain external data connections. | Create a summary report of sheets that need refresh because they host QueryTables. | Automate cleanup by locating worksheets with QueryTables for removal or update.
// AI Prompts: Generate a C# method that returns a List<string> of worksheet names containing QueryTables with Aspose.Cells. | Show how to modify the utility to also display the count of QueryTables per worksheet. | Explain how to add comprehensive error handling for missing files, load failures, and empty QueryTables collections.

using System;
using Aspose.Cells;

namespace AsposeCellsQueryTableUtility
{
    // A command‑line program that loads an Excel workbook with Aspose.Cells, scans each worksheet, and prints the names of those that contain one or more QueryTables.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect the Excel file path as the first command‑line argument
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsQueryTableUtility <path-to-workbook>");
                return;
            }

            string workbookPath = args[0];

            // Load the workbook (uses the standard Aspose.Cells load mechanism)
            Workbook workbook = new Workbook(workbookPath);

            // Iterate over all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // If the worksheet has one or more query tables, output its name
                if (sheet.QueryTables.Count > 0)
                {
                    Console.WriteLine($"Worksheet containing QueryTables: {sheet.Name}");
                }
            }
        }
    }
}
