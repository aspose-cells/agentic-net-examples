// Title: Extract Pivot Cache with LightCells in Aspose.Cells (C#) and Export to CSV
// Description: The sample checks for the source file, loads the workbook with LightCells and pivot‑cache parsing enabled, refreshes all pivot tables, starts an access cache for fast read‑only access, walks through each worksheet and its pivot tables to collect row, column and data field items, closes the cache, and writes the gathered information to a CSV report.
// Keywords: Aspose.Cells | LightCells API | C# | pivot cache extraction | ParsingPivotCachedRecords | RefreshPivotTables | AccessCacheOptions | Export pivot data to CSV | large workbook performance | PivotTable items enumeration
// Common Searches: Aspose.Cells load workbook with LightCells and read pivot cache | Export pivot table cache to CSV using C# | How to use AccessCacheOptions for fast pivot cache extraction | Refresh all pivot tables before extracting cache Aspose.Cells | ParsingPivotCachedRecords example
// Developer Intent: Read a workbook’s pivot cache via LightCells and save the details as a CSV file.
// Use Cases: Generate a comprehensive pivot‑cache dump for downstream BI tools. | Ensure pivot tables are up‑to‑date before exporting cache data. | Accelerate extraction from very large workbooks by using AccessCacheOptions.All.
// AI Prompts: Create C# code that extracts the pivot cache with LightCells and outputs JSON instead of CSV. | Show how to include filter field items when enumerating pivot cache records. | Explain how to modify LoadOptions to open password‑protected Excel files while still parsing pivot caches.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample checks for the source file, loads the workbook with LightCells and pivot‑cache parsing enabled, refreshes all pivot tables, starts an access cache for fast read‑only access, walks through each worksheet and its pivot tables to collect row, column and data field items, closes the cache, and writes the gathered information to a CSV report.
class Program
{
    static void Main()
    {
        // Path to the source workbook that contains pivot tables
        string sourceFile = "PivotWorkbook.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourceFile))
        {
            Console.WriteLine($"Source file '{sourceFile}' not found. Please ensure the file exists in the application directory.");
            return;
        }

        try
        {
            // LoadOptions with LightCells enabled and pivot cache parsing turned on
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                ParsingPivotCachedRecords = true // ensures pivot cache is loaded
            };

            // Load the workbook using the LightCells API (the constructor internally uses LightCells when the option is set)
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Refresh all pivot tables to make sure the cache reflects the latest source data
            workbook.Worksheets.RefreshPivotTables();

            // Start access cache for faster read‑only operations (recommended when extracting large amounts of data)
            workbook.StartAccessCache(AccessCacheOptions.All);

            // Prepare a DataTable that will hold the extracted pivot cache information
            DataTable reportTable = new DataTable("PivotCacheReport");
            reportTable.Columns.Add("Worksheet", typeof(string));
            reportTable.Columns.Add("PivotTable", typeof(string));
            reportTable.Columns.Add("FieldType", typeof(string));
            reportTable.Columns.Add("FieldName", typeof(string));
            reportTable.Columns.Add("ItemName", typeof(string));

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables in the current worksheet
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // ----- Row Fields -----
                    foreach (PivotField rowField in pivot.RowFields)
                    {
                        foreach (PivotItem item in rowField.PivotItems)
                        {
                            reportTable.Rows.Add(sheet.Name, pivot.Name, "Row", rowField.Name, item.Name);
                        }
                    }

                    // ----- Column Fields -----
                    foreach (PivotField colField in pivot.ColumnFields)
                    {
                        foreach (PivotItem item in colField.PivotItems)
                        {
                            reportTable.Rows.Add(sheet.Name, pivot.Name, "Column", colField.Name, item.Name);
                        }
                    }

                    // ----- Data Fields (no items, just field names) -----
                    foreach (PivotField dataField in pivot.DataFields)
                    {
                        reportTable.Rows.Add(sheet.Name, pivot.Name, "Data", dataField.Name, string.Empty);
                    }
                }
            }

            // Close the access cache now that extraction is finished
            workbook.CloseAccessCache(AccessCacheOptions.All);

            // Output the extracted data to console (could be written to CSV, DB, etc.)
            Console.WriteLine("Extracted Pivot Cache Data:");
            foreach (DataRow row in reportTable.Rows)
            {
                Console.WriteLine($"{row["Worksheet"]}\t{row["PivotTable"]}\t{row["FieldType"]}\t{row["FieldName"]}\t{row["ItemName"]}");
            }

            // Save the report as a CSV file for external reporting
            string csvPath = "PivotCacheReport.csv";
            using (var writer = new StreamWriter(csvPath))
            {
                // Write header
                writer.WriteLine("Worksheet,PivotTable,FieldType,FieldName,ItemName");
                // Write rows
                foreach (DataRow row in reportTable.Rows)
                {
                    writer.WriteLine($"{row["Worksheet"]},{row["PivotTable"]},{row["FieldType"]},{row["FieldName"]},{row["ItemName"]}");
                }
            }

            Console.WriteLine($"Report saved to {csvPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
