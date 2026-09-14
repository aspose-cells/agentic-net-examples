// Title: Identify and list worksheet names that contain QueryTables using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that opens a workbook, scans each worksheet, and prints the names of worksheets where the QueryTables collection is not empty. | Create a reusable method in C# that receives an Excel file path and returns a List<string> of worksheet names that have at least one QueryTable, using Aspose.Cells. | Adapt the example to write the identified worksheet names to a text file instead of the console, while still detecting QueryTables with Aspose.Cells.
// Common Searches: aspnet c# how to get worksheet names that contain query tables with Aspose.Cells | list Excel sheets with data connections using Aspose.Cells .NET | retrieve names of worksheets that have query tables in a workbook via Aspose.Cells API | C# Aspose.Cells enumerate worksheets with QueryTable objects | find worksheets with external data connections in an Excel file using Aspose.Cells
// Tags: Aspose.Cells detect QueryTables in worksheets | C# enumerate worksheets with data connections | list worksheet names containing QueryTable objects | Aspose.Cells workbook query table detection | retrieve worksheet names for external data sources .NET

using System;
using Aspose.Cells;

// The utility loads an Excel workbook supplied via command line, iterates through all worksheets, checks each sheet's QueryTables collection, outputs the names of sheets that contain at least one QueryTable, and saves the workbook unchanged.
class Program
{
    static void Main(string[] args)
    {
        // Ensure a file path is provided
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: Program <excel-file-path>");
            return;
        }

        string filePath = args[0];

        // Load the workbook (load rule)
        Workbook workbook = new Workbook(filePath);

        // Iterate over all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check if the worksheet contains any query tables
            if (sheet.QueryTables.Count > 0)
            {
                // Output the name of the worksheet that has query tables
                Console.WriteLine($"Worksheet with QueryTable: {sheet.Name}");
            }
        }

        // Save the workbook (save rule) – no modifications are made, but rule compliance is kept
        workbook.Save(filePath);
    }
}
