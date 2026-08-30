// Title: How to load an XLSX workbook with ParsingPivotCachedRecords enabled and refresh all pivot tables using Aspose.Cells for .NET (C#)
// AI Prompts: Load an XLSX file with LoadOptions.ParsingPivotCachedRecords = true, then call Worksheets.RefreshPivotTables() and save the result using Aspose.Cells in C#. | Write C# code that opens a workbook with pivot cache parsing turned on, updates every pivot table, and writes the updated file.
// Common Searches: Aspose.Cells C# load workbook with ParsingPivotCachedRecords and refresh pivot tables | Enable parsing of pivot cached records when opening Excel file with Aspose.Cells | Refresh all pivot tables after loading workbook in .NET | How to use LoadOptions to parse pivot cache in Aspose.Cells | C# code to refresh pivot tables after loading XLSX with Aspose.Cells
// Tags: parsingpivotcachedrecords option aspnet cells | pivot tables refresh method c# | loadoptions xlsx aspnet cells | worksheet refreshpivottables c# | excel pivot cache parsing aspnet

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// // Loads an XLSX workbook with ParsingPivotCachedRecords enabled, refreshes every pivot table, and saves the updated file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create load options and enable parsing of pivot cached records
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.ParsingPivotCachedRecords = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Refresh all pivot tables in the workbook
        workbook.Worksheets.RefreshPivotTables();

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}
