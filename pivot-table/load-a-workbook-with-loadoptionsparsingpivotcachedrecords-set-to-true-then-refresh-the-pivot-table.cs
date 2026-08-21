// Title: Refresh Pivot Tables After Loading Workbook with ParsingPivotCachedRecords – Aspose.Cells for .NET
// Description: Demonstrates how to enable the ParsingPivotCachedRecords option in LoadOptions, load an XLSX workbook, refresh all pivot tables programmatically, and save the updated file using Aspose.Cells for C#.
// Keywords: Aspose.Cells | LoadOptions | ParsingPivotCachedRecords | refresh pivot tables | C# | Excel automation | pivot cache parsing | Workbook.RefreshPivotTables
// Common Searches: Aspose.Cells enable ParsingPivotCachedRecords | refresh all pivot tables after loading workbook C# | load Excel file with pivot cache parsing Aspose | how to programmatically refresh pivot tables .NET | Aspose.Cells LoadOptions pivot cache example
// Developer Intent: Load an XLSX workbook with pivot cache parsing turned on and refresh its pivot tables using Aspose.Cells.
// Use Cases: Automate batch processing of Excel files that contain pivot tables, ensuring cached data is parsed and tables are refreshed before saving. | Refresh pivot tables after external data changes when re‑opening a workbook with the ParsingPivotCachedRecords flag enabled. | Integrate pivot‑table refresh into a server‑side reporting pipeline that generates up‑to‑date Excel reports.
// AI Prompts: Generate C# code that loads an XLSX file with ParsingPivotCachedRecords set to true, refreshes all pivot tables, and saves the workbook using Aspose.Cells. | Explain the impact of the ParsingPivotCachedRecords property on pivot table data when a workbook is loaded with Aspose.Cells. | Create a reusable method that accepts a file path, applies LoadOptions with pivot cache parsing, refreshes pivot tables, and returns the modified Workbook object.

using System;
using Aspose.Cells;

// Demonstrates how to enable the ParsingPivotCachedRecords option in LoadOptions, load an XLSX workbook, refresh all pivot tables programmatically, and save the updated file using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Configure load options to parse pivot cached records
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
