// Title: Load an Excel template in streaming mode and process smart markers on a specific worksheet with Aspose.Cells for .NET
// Common Searches: how to load an Excel file in low‑memory mode with Aspose.Cells C# | process smart markers on a single worksheet using Aspose.Cells streaming | Aspose.Cells memory‑efficient smart marker processing example | C# bind DataTable to WorkbookDesigner for smart marker replacement | using LoadOptions MemorySetting.MemoryPreference for large Excel workbooks
// Tags: streaming loadoptions memorypreference | process smart markers single worksheet | WorkbookDesigner setdatasource datatable | low memory excel processing csharp | save workbook xlsx Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel template with smart markers in streaming (low‑memory) mode by setting MemorySetting.MemoryPreference in LoadOptions, binding a DataTable to a WorkbookDesigner, processing smart markers only on the first worksheet while preserving unknown markers, and saving the result as an XLSX file, enabling memory‑optimized handling of large workbooks.
class Program
{
    static void Main()
    {
        // --------------------------------------------------------------------
        // 1. Define file paths (adjust these paths to your environment)
        // --------------------------------------------------------------------
        string templatePath = "SmartMarkerTemplate.xlsx";   // workbook that contains smart markers
        string outputPath   = "ProcessedOutput.xlsx";      // result file

        // Verify that the template file exists to avoid FileNotFoundException
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        // --------------------------------------------------------------------
        // 2. Prepare a simple data source for the smart markers
        // --------------------------------------------------------------------
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("Name",   typeof(string));
        employees.Columns.Add("Age",    typeof(int));
        employees.Columns.Add("Salary", typeof(double));

        employees.Rows.Add("John Doe", 30, 50000);
        employees.Rows.Add("Jane Smith", 28, 55000);
        employees.Rows.Add("Bob Johnson", 35, 62000);

        // --------------------------------------------------------------------
        // 3. Configure LoadOptions for streaming (memory‑optimized) mode
        // --------------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            // Use streaming / low‑memory mode
            MemorySetting = MemorySetting.MemoryPreference
        };

        // NOTE: LoadFilter is omitted because the current Aspose.Cells version
        // does not expose the required properties. The workbook will be loaded
        // in streaming mode, which still reduces memory consumption.

        // --------------------------------------------------------------------
        // 4. Load the workbook with the above options
        // --------------------------------------------------------------------
        Workbook workbook;
        try
        {
            workbook = new Workbook(templatePath, loadOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // --------------------------------------------------------------------
        // 5. Create a WorkbookDesigner, bind the data source and process
        //    only the first worksheet (index 0). The 'true' flag preserves
        //    any unrecognized smart markers.
        // --------------------------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(employees);

        try
        {
            // Process smart markers on the first worksheet only
            designer.Process(0, true);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing smart markers: {ex.Message}");
            return;
        }

        // --------------------------------------------------------------------
        // 6. Save the processed workbook
        // --------------------------------------------------------------------
        try
        {
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Processing complete. File saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }

        // --------------------------------------------------------------------
        // OPTIONAL: If the template contains additional large sections,
        // you can repeat steps 3‑5 with different LoadOptions ranges,
        // processing each part sequentially to keep memory usage low.
        // --------------------------------------------------------------------
    }
}
