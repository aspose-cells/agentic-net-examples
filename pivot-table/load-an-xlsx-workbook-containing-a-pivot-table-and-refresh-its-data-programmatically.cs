// Title: How to refresh all pivot tables in an XLSX workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing .xlsx file, calls Workbook.Worksheets.RefreshPivotTables(), and saves the result to a new file. | Show a step‑by‑step example of using Aspose.Cells to programmatically refresh every pivot table in a workbook and write the updated workbook to disk. | Provide a minimal console application that demonstrates loading a workbook, refreshing its pivot tables, and persisting the changes with Aspose.Cells.
// Common Searches: aspnet refresh all pivot tables in existing excel file using aspose.cells | c# code to refresh pivot tables in an xlsx workbook with aspose.cells library | how to programmatically update pivot cache after data change using aspose.cells | example of RefreshPivotTables method in Aspose.Cells for .NET | refresh pivot tables in workbook and save as new file using C# Aspose.Cells
// Tags: Aspose.Cells pivot table refresh API | C# load XLSX and refresh pivot caches | programmatic Excel pivot refresh .NET | save workbook after pivot update Aspose.Cells | update all pivot tables in workbook

using System;
using Aspose.Cells;

namespace AsposeCellsPivotRefresh
{
    // // This program loads an existing XLSX workbook, refreshes all pivot tables using the RefreshPivotTables method, and saves the refreshed workbook to a new file.
    public class Program
    {
        public static void Main()
        {
            // Load the existing workbook that contains a pivot table
            Workbook workbook = new Workbook("input.xlsx");

            // Refresh all pivot tables in the workbook
            workbook.Worksheets.RefreshPivotTables();

            // Save the workbook after refreshing the pivot tables
            workbook.Save("output.xlsx");
        }
    }
}
