// Title: C# script to delete every pivot table in all worksheets of an Excel workbook with Aspose.Cells
// Common Searches: aspnet delete all pivot tables from an Excel workbook using Aspose.Cells | c# code to clear pivot tables on each sheet with Aspose.Cells | remove pivot tables from every worksheet programmatically Aspose.Cells .NET | how to iterate worksheets and delete pivot tables using Aspose.Cells C# | batch process Excel file to strip all pivot tables with Aspose.Cells
// Tags: Aspose.Cells clear PivotTables collection | C# loop through worksheets Aspose.Cells | delete workbook-wide pivot tables Aspose.Cells | strip Excel pivot tables using Aspose.Cells | programmatic pivot table removal C#

using Aspose.Cells;
using System;

// Loads an Excel file, iterates through each worksheet, clears the PivotTables collection on every sheet, and saves the workbook, effectively removing all pivot tables using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Remove all pivot tables from the current worksheet
            sheet.PivotTables.Clear();
        }

        // Save the workbook after pivot tables have been removed
        workbook.Save("output.xlsx");
    }
}
