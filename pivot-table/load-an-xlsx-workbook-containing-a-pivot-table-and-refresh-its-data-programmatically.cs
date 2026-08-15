// Title: Refresh All PivotTables in an XLSX Workbook with Aspose.Cells for .NET
// Description: Loads an existing XLSX file, programmatically refreshes every PivotTable using Aspose.Cells, and saves the updated workbook to a new location.
// Keywords: Aspose.Cells | .NET | C# | refresh pivot tables | Workbook.RefreshPivotTables | programmatic Excel pivot update | pivot cache refresh
// Common Searches: how to refresh pivot tables in Excel using Aspose.Cells C# | Aspose.Cells refresh all PivotTables in a workbook | C# code to update pivot cache programmatically | refresh Excel pivot tables on server side | Aspose.Cells RefreshPivotTables method example
// Developer Intent: Update all PivotTables in an existing XLSX workbook and write the refreshed file.
// Use Cases: Automatically refresh pivot reports after nightly data imports before distribution. | Integrate pivot table refresh into a web service that generates Excel analytics on demand. | Maintain up‑to‑date PivotTables in batch‑processed workbooks for regulatory reporting.
// AI Prompts: Write C# code that refreshes only specific PivotTables in a workbook using Aspose.Cells. | Explain the differences between RefreshPivotTables() and manual cache refresh for large data sets. | Suggest performance tips for refreshing PivotTables when the source range contains millions of rows.

using System;
using Aspose.Cells;

// Loads an existing XLSX file, programmatically refreshes every PivotTable using Aspose.Cells, and saves the updated workbook to a new location.
class Program
{
    static void Main()
    {
        // Path to the source workbook that contains a pivot table
        string inputPath = "input.xlsx";

        // Path where the refreshed workbook will be saved
        string outputPath = "output.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // Refresh all PivotTables in the workbook
        workbook.Worksheets.RefreshPivotTables();

        // Save the updated workbook
        workbook.Save(outputPath);
    }
}
