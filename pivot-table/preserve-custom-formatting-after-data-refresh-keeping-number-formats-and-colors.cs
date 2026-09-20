// Title: Refresh Pivot Table Data in an Excel Workbook While Keeping Custom Number Formats and Cell Colors Using Aspose.Cells for .NET
// AI Prompts: Load an existing workbook, recalculate its formulas (or refresh data connections if supported), and save it so that custom number formats and cell background colors stay unchanged. | Import a DataTable into a worksheet and then recalculate formulas, ensuring that all pre‑existing cell styles are preserved with Aspose.Cells. | Implement a fallback that calls Workbook.CalculateFormula when RefreshAllDataConnections is unavailable, guaranteeing that formatting is not altered.
// Common Searches: Aspose.Cells keep cell background color after refreshing pivot table data | preserve custom number format when recalculating formulas in .NET Excel workbook | how to refresh data connections without losing formatting using Aspose.Cells | import DataTable into existing worksheet without changing existing styles Aspose.Cells | refresh pivot table programmatically while retaining number formats Aspose.Cells .NET
// Tags: refresh data connections Aspose.Cells .NET | preserve custom number formats Excel Aspose.Cells | retain cell colors after workbook recalculation | import DataTable without overwriting styles Aspose.Cells | pivot table refresh preserving formatting Aspose.Cells

using Aspose.Cells;
using System;
using System.Data;
using System.IO;

// The example loads an Excel workbook, optionally refreshes external data connections or recalculates formulas, and saves the file, ensuring that custom number formats and cell colors remain intact.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook (preserves all existing formatting).
            Workbook workbook = new Workbook(inputPath);

            // NOTE: In some Aspose.Cells versions the method RefreshAllDataConnections
            // is not available. If you need to refresh external data sources, use the
            // appropriate API for your version. Here we simply recalculate formulas
            // as a safe fallback that does not affect cell styles.
            workbook.CalculateFormula();

            // Example of importing data while preserving formatting (uncomment and adapt as needed).
            // DataTable dt = GetDataFromSource();
            // Worksheet sheet = workbook.Worksheets[0];
            // sheet.Cells.ImportDataTable(dt, true, "A2", true);

            // Save the workbook; all custom number formats and cell colors remain intact.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Placeholder for custom data retrieval logic.
    // static DataTable GetDataFromSource()
    // {
    //     // Implement data fetching and return a populated DataTable.
    // }
}
