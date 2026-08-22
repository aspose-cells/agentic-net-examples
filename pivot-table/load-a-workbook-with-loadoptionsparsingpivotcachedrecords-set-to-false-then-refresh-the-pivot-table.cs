// Title: Refresh all pivot tables in an XLSX workbook after disabling pivot cache parsing with Aspose.Cells for .NET
// AI Prompts: Open an XLSX file using Aspose.Cells LoadOptions with ParsingPivotCachedRecords set to false, refresh every pivot table, and save the result. | Load a workbook without parsing pivot cached records, invoke the RefreshPivotTables method programmatically, and export the updated file using C#.
// Common Searches: Aspose.Cells how to disable pivot cache parsing and refresh pivots in C# | load xlsx without pivot cached records then refresh pivot tables using .NET | refresh all pivot tables after opening workbook with ParsingPivotCachedRecords false | C# Aspose.Cells load workbook with custom LoadOptions and update pivot tables
// Tags: disable pivot cache parsing Aspose.Cells | update pivot tables after workbook load Aspose.Cells | LoadOptions ParsingPivotCachedRecords false example | load workbook with custom LoadOptions Aspose.Cells | Aspose.Cells C# programmatic pivot update

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// // Load an XLSX workbook with pivot cache parsing turned off, refresh all pivot tables, and save the updated file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create load options and disable parsing of pivot cached records
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.ParsingPivotCachedRecords = false;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Refresh all pivot tables in the workbook
        workbook.Worksheets.RefreshPivotTables();

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}
