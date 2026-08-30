// Title: C# script to delete every pivot table in all worksheets of an Excel workbook with Aspose.Cells
// AI Prompts: Generate a C# console application that loads an .xlsx file, removes all pivot tables from every worksheet using Workbook.Worksheets.ClearPivottables, and writes the result to a new file. | Show how to use Aspose.Cells to batch‑clear pivot tables across a workbook without iterating each sheet manually. | Provide code that demonstrates removing pivot tables from a workbook and saving it, suitable for inclusion in an automated build or deployment script.
// Common Searches: aspnet remove all pivot tables from an Excel workbook using Aspose.Cells | C# clear pivot tables in every sheet programmatically Aspose.Cells | how to use ClearPivottables method in Aspose.Cells .NET | batch process to delete pivot tables from multiple worksheets with Aspose.Cells | automate pivot table cleanup in Excel files using C# Aspose.Cells
// Tags: Aspose.Cells ClearPivottables method | remove pivot tables workbook C# | batch delete Excel pivot tables Aspose.Cells | clear all pivot tables .NET | automated pivot table cleanup Aspose.Cells

using System;
using Aspose.Cells;

// // Loads an Excel workbook, calls Workbook.Worksheets.ClearPivottables to delete every pivot table across all worksheets, and saves the modified file.
class RemoveAllPivotTables
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Clear all pivot tables from every worksheet in the workbook
        workbook.Worksheets.ClearPivottables();

        // Save the workbook after removal
        workbook.Save("output.xlsx");
    }
}
