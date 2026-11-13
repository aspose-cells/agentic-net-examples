// Title: How to list worksheet names that contain query tables using Aspose.Cells for .NET (C#)
// Common Searches: C# code to get names of Excel worksheets that have query tables with Aspose.Cells | Aspose.Cells .NET find worksheets with data connections (query tables) in a workbook | enumerate sheets containing query tables in an Excel file using Aspose.Cells library | retrieve worksheet names where the QueryTables collection is not empty with Aspose.Cells C#
// Tags: Aspose.Cells worksheet query table enumeration | C# detect Excel sheets with query tables | Aspose.Cells query tables per worksheet | list Excel worksheets having data connections | iterate workbook to find query tables Aspose

using System;
using Aspose.Cells;

// The example loads an Excel workbook via Aspose.Cells, iterates through each worksheet, checks if the QueryTables collection has entries, and prints the names of worksheets that contain one or more query tables.
class QueryTableWorksheetLister
{
    static void Main(string[] args)
    {
        // Ensure a file path is provided
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: QueryTableWorksheetLister <excel-file-path>");
            return;
        }

        string filePath = args[0];

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check if the worksheet contains any query tables
            if (sheet.QueryTables != null && sheet.QueryTables.Count > 0)
            {
                // Output the name of the worksheet
                Console.WriteLine(sheet.Name);
            }
        }
    }
}
