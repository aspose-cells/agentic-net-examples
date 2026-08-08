// Title: Parallel processing of multiple Excel templates with smart markers using Aspose.Cells for .NET
// Description: Loads several Excel templates that contain smart markers, creates a separate WorkbookDesigner for each, binds individual DataTables, enables MultiThreadReading, processes the templates concurrently with Parallel.For, and merges the resulting workbooks into a single file (MergedResult.xlsx).
// Keywords: Aspose.Cells parallel processing | WorkbookDesigner multi‑thread | smart markers concurrent C# | combine multiple workbooks Aspose | MultiThreadReading cells | Parallel.For Excel generation
// Common Searches: Aspose.Cells process smart markers in parallel | C# merge workbooks after parallel processing | Enable MultiThreadReading for WorkbookDesigner | Parallel.For Aspose.Cells example | Combine multiple template workbooks .NET
// Developer Intent: Run separate WorkbookDesigner instances on different templates simultaneously and consolidate the outputs into one workbook.
// Use Cases: Generate a master report by populating several smart‑marker templates with distinct data sets in parallel, then merging them. | Speed up bulk mail‑merge style Excel creation by assigning each template to its own thread and combining the results. | Aggregate departmental spreadsheets processed concurrently into a single master workbook to reduce overall runtime.
// AI Prompts: Provide C# code that creates a WorkbookDesigner for each Excel template, processes them inside Parallel.For, and merges the workbooks with Aspose.Cells. | Explain how to safely enable MultiThreadReading on worksheets when using Parallel.For with smart markers. | Suggest best practices for error handling and logging in a parallel Aspose.Cells workbook processing scenario.

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// Loads several Excel templates that contain smart markers, creates a separate WorkbookDesigner for each, binds individual DataTables, enables MultiThreadReading, processes the templates concurrently with Parallel.For, and merges the resulting workbooks into a single file (MergedResult.xlsx).
class MultiThreadWorkbookDesignerDemo
{
    static void Main()
    {
        // Paths to template workbooks (each contains smart markers)
        string[] templates = { "Template1.xlsx", "Template2.xlsx", "Template3.xlsx" };

        // Prepare a simple data source for each template (DataTable used as example)
        List<DataTable> dataSources = new List<DataTable>();
        for (int i = 0; i < templates.Length; i++)
        {
            DataTable dt = new DataTable("Table" + i);
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Value", typeof(int));
            dt.Rows.Add("ItemA", i * 10);
            dt.Rows.Add("ItemB", i * 20);
            dataSources.Add(dt);
        }

        // Array to hold the processed workbooks from each thread
        Workbook[] processedWorkbooks = new Workbook[templates.Length];

        // Process each template concurrently
        Parallel.For(0, templates.Length, index =>
        {
            try
            {
                string templatePath = templates[index];

                // Verify that the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}. Skipping this entry.");
                    return;
                }

                // Load the template workbook
                Workbook wb = new Workbook(templatePath);

                // Enable multi‑thread reading for the cells collection (required for safe concurrent reads)
                wb.Worksheets[0].Cells.MultiThreadReading = true;

                // Create a WorkbookDesigner bound to this workbook
                WorkbookDesigner designer = new WorkbookDesigner(wb);

                // Bind the data source to the smart marker name "Data"
                designer.SetDataSource("Data", dataSources[index]);

                // Process the smart markers and populate the workbook
                designer.Process();

                // Store the processed workbook for later merging
                processedWorkbooks[index] = designer.Workbook;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing template '{templates[index]}': {ex.Message}");
            }
        });

        // Create an empty workbook that will hold the merged result
        Workbook finalWorkbook = new Workbook();

        // Remove the default empty sheet created by the constructor, if present
        if (finalWorkbook.Worksheets.Count > 0)
        {
            finalWorkbook.Worksheets.RemoveAt(0);
        }

        // Merge each processed workbook into the final workbook
        foreach (Workbook wb in processedWorkbooks)
        {
            if (wb != null)
            {
                finalWorkbook.Combine(wb);
            }
        }

        // Save the merged workbook to disk
        try
        {
            finalWorkbook.Save("MergedResult.xlsx");
            Console.WriteLine("Merged workbook saved as 'MergedResult.xlsx'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save merged workbook: {ex.Message}");
        }
    }
}
