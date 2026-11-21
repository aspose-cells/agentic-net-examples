// Title: Parallel processing of Aspose.Cells smart markers with WorkbookDesigner to generate multiple Excel files from massive DataTables (C#)
// Common Searches: how to process Aspose.Cells smart markers in parallel using C# | parallel smart marker population with WorkbookDesigner for large DataTables | generate separate Excel workbooks from massive data sets using Aspose.Cells multithreading | C# Parallel.ForEach example for Aspose.Cells smart marker templates | thread‑safe way to use WorkbookDesigner for each dataset in Aspose.Cells
// Tags: parallel smart marker processing Aspose.Cells | WorkbookDesigner thread-safe execution | massive DataTable to Excel conversion Aspose.Cells | C# multithreaded Excel generation with smart markers | high‑performance smart marker population

using System;
using System.Data;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample loads a template workbook, creates a list of large DataTables, and uses Parallel.ForEach to instantiate a fresh WorkbookDesigner for each dataset. It assigns the data source, runs smart marker processing, and saves each populated workbook to a uniquely indexed Excel file, handling errors per iteration.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the template workbook that contains smart markers.
            const string templatePath = "Template.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Prepare a collection of massive data sets to be merged.
            List<DataTable> dataSets = GetMassiveDataSets();

            // Process each data set in parallel to accelerate population.
            Parallel.ForEach(dataSets, (dataSet, state, index) =>
            {
                try
                {
                    // Load a fresh copy of the template for each thread.
                    Workbook wb = new Workbook(templatePath);

                    // Initialize the workbook designer for smart marker processing.
                    WorkbookDesigner designer = new WorkbookDesigner(wb);

                    // Assign the data source; the name "Data" must match the smart marker tags in the template.
                    designer.SetDataSource("Data", dataSet);

                    // Execute smart marker processing.
                    designer.Process();

                    // Save the populated workbook; each file gets a unique index.
                    string outputPath = $"Result_{index}.xlsx";
                    wb.Save(outputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing dataset {index}: {ex.Message}");
                }
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Generates a list of large DataTables to simulate massive data sets.
    static List<DataTable> GetMassiveDataSets()
    {
        var list = new List<DataTable>();

        for (int i = 0; i < 10; i++) // Example: 10 separate data sets.
        {
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Value", typeof(double));

            // Populate each table with a large number of rows.
            for (int r = 0; r < 100_000; r++)
            {
                dt.Rows.Add($"Item_{i}_{r}", r * 0.1);
            }

            list.Add(dt);
        }

        return list;
    }
}
