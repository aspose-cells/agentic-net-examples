// Title: Load Workbook with LightCells, Enable Pivot Cache Parsing, Refresh and Extract Pivot Data (Aspose.Cells C#)
// Description: C# example that loads an XLSX workbook using Aspose.Cells LightCells API with ParsingPivotCachedRecords enabled, refreshes all pivot tables, iterates through each pivot table, and prints cached row, column, and data field items. The refreshed workbook can be saved for downstream reporting.
// Keywords: Aspose.Cells | C# | LoadOptions | ParsingPivotCachedRecords | LightCells API | pivot cache extraction | refresh pivot tables programmatically | read pivot field items | external reporting | save refreshed workbook
// Common Searches: how to enable pivot cache parsing with Aspose.Cells | C# read cached pivot table data Aspose.Cells | refresh all pivot tables before extracting cache | extract row and column items from pivot tables using Aspose.Cells | save workbook after pivot refresh Aspose.Cells
// Developer Intent: Load a workbook, refresh its pivot tables, and retrieve cached pivot field values for reporting or analysis.
// Use Cases: Generate external reports by pulling cached pivot items without connecting to the original data source. | Validate pivot table structure and cached calculations after a data refresh. | Create a refreshed copy of the workbook for downstream processing or archival.
// AI Prompts: Show a C# snippet that loads an XLSX with Aspose.Cells, enables ParsingPivotCachedRecords, refreshes pivot tables, and enumerates row, column, and data field items. | Explain why refreshing pivot tables before reading cached values is necessary and how Aspose.Cells implements it. | Provide code to export the extracted pivot cache items to CSV or JSON instead of writing to the console.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# example that loads an XLSX workbook using Aspose.Cells LightCells API with ParsingPivotCachedRecords enabled, refreshes all pivot tables, iterates through each pivot table, and prints cached row, column, and data field items. The refreshed workbook can be saved for downstream reporting.
class ExtractPivotCache
{
    static void Main()
    {
        // Path to the workbook that contains pivot tables
        string inputPath = "PivotData.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Enable parsing of pivot cached records while loading the file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                ParsingPivotCachedRecords = true
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Refresh all pivot tables to ensure the cache reflects the latest source data
            workbook.Worksheets.RefreshPivotTables();

            // Iterate through each worksheet that contains pivot tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.PivotTables.Count == 0) continue;

                Console.WriteLine($"Worksheet: {sheet.Name}");

                // Iterate through each pivot table in the worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    Console.WriteLine($"  PivotTable: {pivotTable.Name}");

                    // ----- Row Fields -----
                    foreach (PivotField rowField in pivotTable.RowFields)
                    {
                        Console.WriteLine($"    Row Field: {rowField.Name}");
                        foreach (PivotItem item in rowField.PivotItems)
                        {
                            Console.WriteLine($"      Item: {item.Value}");
                        }
                    }

                    // ----- Column Fields -----
                    foreach (PivotField colField in pivotTable.ColumnFields)
                    {
                        Console.WriteLine($"    Column Field: {colField.Name}");
                        foreach (PivotItem item in colField.PivotItems)
                        {
                            Console.WriteLine($"      Item: {item.Value}");
                        }
                    }

                    // ----- Data Fields (cached values) -----
                    foreach (PivotField dataField in pivotTable.DataFields)
                    {
                        Console.WriteLine($"    Data Field: {dataField.Name}");
                        foreach (PivotItem item in dataField.PivotItems)
                        {
                            // The cached numeric value is stored in the Value property
                            Console.WriteLine($"      Item: {item.Value}");
                        }
                    }
                }
            }

            // Save the workbook after refresh (optional)
            string outputPath = "PivotData_Refreshed.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
