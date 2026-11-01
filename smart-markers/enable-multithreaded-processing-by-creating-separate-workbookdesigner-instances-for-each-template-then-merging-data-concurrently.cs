// Title: How to process multiple Excel templates concurrently with WorkbookDesigner and merge them into a single workbook using Aspose.Cells for .NET
// Common Searches: c# process multiple Excel templates in parallel using Aspose.Cells WorkbookDesigner | merge worksheets from several processed workbooks into one file with Aspose.Cells | thread‑safe usage of WorkbookDesigner inside Parallel.For loop | create master workbook from multiple smart‑marker templates Aspose.Cells .NET | parallel smart marker processing with separate WorkbookDesigner instances
// Tags: parallel WorkbookDesigner processing Aspose.Cells | merge multiple workbooks Aspose.Cells .NET | smart markers multithreaded execution C# | copy worksheets to master workbook Aspose.Cells | thread‑safe Excel template processing Aspose.Cells

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads three Excel template files, assigns a distinct DataTable to each, and processes them concurrently using separate WorkbookDesigner instances inside a Parallel.For loop. Processed workbooks are collected, their worksheets are copied into a new master workbook, and the combined result is saved as MergedResult.xlsx. The code includes file existence checks, error handling, and creates the output directory if needed.
class MultiThreadedWorkbookDesigner
{
    static void Main()
    {
        try
        {
            // Paths to the template files
            var templatePaths = new List<string>
            {
                @"Templates\Template1.xlsx",
                @"Templates\Template2.xlsx",
                @"Templates\Template3.xlsx"
            };

            // Verify that all template files exist
            foreach (var path in templatePaths)
            {
                if (!File.Exists(path))
                {
                    throw new FileNotFoundException($"Template file not found: {path}");
                }
            }

            // Prepare data sources for each template (example using DataTable)
            var dataSources = new List<object>
            {
                GetSampleDataTable("Employees"),
                GetSampleDataTable("Products"),
                GetSampleDataTable("Orders")
            };

            // Array to hold the processed workbooks
            var processedWorkbooks = new Workbook[templatePaths.Count];

            // Process each template in parallel using separate WorkbookDesigner instances
            Parallel.For(0, templatePaths.Count, i =>
            {
                try
                {
                    var designer = new WorkbookDesigner();

                    // Load the template workbook
                    designer.Workbook = new Workbook(templatePaths[i]);

                    // Set the data source (the name "Data" is used in the template)
                    designer.SetDataSource("Data", dataSources[i]);

                    // Process the template (merge data into the workbook)
                    designer.Process();

                    // Store the resulting workbook
                    processedWorkbooks[i] = designer.Workbook;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing template '{templatePaths[i]}': {ex.Message}");
                    throw;
                }
            });

            // Merge all processed workbooks into a single workbook
            var masterWorkbook = new Workbook();

            // Remove the default empty sheet created by the master workbook
            if (masterWorkbook.Worksheets.Count > 0)
            {
                masterWorkbook.Worksheets.RemoveAt(0);
            }

            foreach (var wb in processedWorkbooks)
            {
                foreach (Worksheet ws in wb.Worksheets)
                {
                    // Copy each worksheet into the master workbook using the sheet name
                    masterWorkbook.Worksheets.AddCopy(ws.Name);
                }
            }

            // Ensure the output directory exists
            var outputDir = @"Output";
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the merged workbook
            masterWorkbook.Save(Path.Combine(outputDir, "MergedResult.xlsx"));
            Console.WriteLine("Merged workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }

    // Helper method to create a sample DataTable (replace with real data retrieval)
    private static DataTable GetSampleDataTable(string tableName)
    {
        var dt = new DataTable(tableName);
        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Value", typeof(double));

        // Populate with dummy rows
        for (int i = 1; i <= 5; i++)
        {
            dt.Rows.Add(i, $"{tableName}_Item{i}", i * 10.5);
        }

        return dt;
    }
}
