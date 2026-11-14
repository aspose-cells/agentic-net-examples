// Title: Generate and merge multiple Excel files in parallel using WorkbookDesigner with Aspose.Cells for .NET
// Common Searches: how to use WorkbookDesigner to process several Excel templates concurrently in C# | parallel smart markers processing with Aspose.Cells .NET | merge worksheets from multiple WorkbookDesigner instances into one workbook | thread‑safe collection for storing processed Aspose.Cells workbooks | run Aspose.Cells template merging inside Parallel.For loop
// Tags: parallel WorkbookDesigner processing Aspose.Cells | multithreaded smart markers template merging | concurrent workbook collection C# | merge worksheets from multiple workbooks Aspose.Cells | thread‑safe Excel file generation .NET

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using Aspose.Cells;

// The example loads several Excel template files, creates a separate WorkbookDesigner for each inside a Parallel.For loop, assigns DataTable sources to the "Data" smart‑marker, processes the templates concurrently, stores the resulting Workbook objects in a thread‑safe ConcurrentBag, then copies all worksheets into a single final workbook and saves it as MergedResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to the template files
            var templates = new List<string>
            {
                "Template1.xlsx",
                "Template2.xlsx",
                "Template3.xlsx"
            };

            // Corresponding data sources for each template
            var dataSources = new List<object>
            {
                GetDataTable1(),
                GetDataTable2(),
                GetDataTable3()
            };

            // Thread‑safe collection to store the processed workbooks
            var processedWorkbooks = new ConcurrentBag<Workbook>();

            // Process each template in parallel, each with its own WorkbookDesigner
            Parallel.For(0, templates.Count, i =>
            {
                try
                {
                    string templatePath = templates[i];

                    // Ensure the template file exists
                    if (!File.Exists(templatePath))
                    {
                        Console.WriteLine($"Template file not found: {templatePath}");
                        return;
                    }

                    // Create a new designer instance for this thread
                    var designer = new WorkbookDesigner();

                    // Load the template workbook
                    designer.Workbook = new Workbook(templatePath);

                    // Assign the data source (named "Data" in the template)
                    designer.SetDataSource("Data", dataSources[i]);

                    // Merge the data into the template
                    designer.Process();

                    // Store the resulting workbook
                    processedWorkbooks.Add(designer.Workbook);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing template '{templates[i]}': {ex.Message}");
                }
            });

            // Create a workbook that will hold all merged worksheets
            var finalWorkbook = new Workbook();

            // Add each processed worksheet to the final workbook
            foreach (var wb in processedWorkbooks)
            {
                foreach (Worksheet sheet in wb.Worksheets)
                {
                    // Add a copy of the worksheet to the final workbook using its name
                    finalWorkbook.Worksheets.AddCopy(sheet.Name);
                }
            }

            // Remove the initial empty sheet if it was never used
            if (finalWorkbook.Worksheets.Count > 0 &&
                finalWorkbook.Worksheets[0].Name == "Sheet1" &&
                finalWorkbook.Worksheets[0].Cells.MaxDataColumn == 0 &&
                finalWorkbook.Worksheets[0].Cells.MaxDataRow == 0)
            {
                finalWorkbook.Worksheets.RemoveAt(0);
            }

            // Save the merged result
            finalWorkbook.Save("MergedResult.xlsx");
            Console.WriteLine("Merged workbook saved as 'MergedResult.xlsx'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Sample data for the first template
    static DataTable GetDataTable1()
    {
        var dt = new DataTable();
        dt.Columns.Add("Name");
        dt.Columns.Add("Value");
        dt.Rows.Add("Item1", 10);
        dt.Rows.Add("Item2", 20);
        return dt;
    }

    // Sample data for the second template
    static DataTable GetDataTable2()
    {
        var dt = new DataTable();
        dt.Columns.Add("Product");
        dt.Columns.Add("Price");
        dt.Rows.Add("A", 5.5);
        dt.Rows.Add("B", 7.25);
        return dt;
    }

    // Sample data for the third template
    static DataTable GetDataTable3()
    {
        var dt = new DataTable();
        dt.Columns.Add("Country");
        dt.Columns.Add("Population");
        dt.Rows.Add("X", 1000000);
        dt.Rows.Add("Y", 2000000);
        return dt;
    }
}
