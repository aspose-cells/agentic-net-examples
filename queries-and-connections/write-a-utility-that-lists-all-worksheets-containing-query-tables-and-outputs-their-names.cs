// Title: C# console utility that lists worksheet names containing query tables using Aspose.Cells
// AI Prompts: Create a C# console application with Aspose.Cells that takes a workbook file path, scans each worksheet for QueryTables, and prints the names of worksheets that have at least one query table. | Generate C# code that loads an Excel workbook via Aspose.Cells, checks the QueryTables collection on every sheet, and outputs the sheet names where query tables are present.
// Common Searches: how to list worksheets that contain query tables with Aspose.Cells in C# | C# Aspose.Cells example to find sheets having query tables | enumerate query tables per worksheet using Aspose.Cells .NET | retrieve names of Excel sheets with query tables via Aspose.Cells | Aspose.Cells C# code to detect worksheets with query tables
// Tags: Aspose.Cells worksheet query table enumeration | C# query table presence check in Excel sheets | Aspose.Cells console utility for query table listing | Excel workbook analysis Aspose.Cells query tables | identify sheets with query tables using Aspose.Cells

using System;
using Aspose.Cells;

// A C# console program that loads an Excel workbook with Aspose.Cells, iterates all worksheets, checks each sheet's QueryTables collection, and writes out the names of worksheets that contain one or more query tables.
class QueryTableWorksheetLister
{
    static void Main(string[] args)
    {
        // Validate input arguments
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: QueryTableWorksheetLister <input workbook path>");
            return;
        }

        string workbookPath = args[0];

        // Load the workbook (Aspose.Cells handles various formats)
        Workbook workbook = new Workbook(workbookPath);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check if the worksheet contains any query tables
            if (sheet.QueryTables != null && sheet.QueryTables.Count > 0)
            {
                // Output the name of the worksheet that has query tables
                Console.WriteLine(sheet.Name);
            }
        }
    }
}
