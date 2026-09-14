// Title: Batch processing multiple Excel templates with individual smart‑marker data sources using Aspose.Cells WorkbookDesigner in C#
// AI Prompts: Generate C# code that loads a list of Excel template files, assigns a distinct DataTable to each WorkbookDesigner, processes smart markers, and writes each workbook to its own output file. | Add error handling so that missing template files are logged and a new blank workbook is created before applying smart markers in the batch loop. | Refactor the example to read template paths, output paths, and JSON‑defined data sources from a configuration file and execute the batch smart‑marker processing.
// Common Searches: Aspose.Cells batch smart marker processing with different data tables in C# | How to use WorkbookDesigner for multiple workbooks in a loop | Create default workbook when template file not found Aspose.Cells | Assign separate DataTable to each smart marker job using Aspose.Cells | C# example for processing several Excel templates with smart markers
// Tags: batch smart marker processing Aspose.Cells | WorkbookDesigner per workbook data source | fallback blank workbook Aspose.Cells | process multiple Excel templates C# | save individual processed workbooks | smart marker data source configuration

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample defines a collection of jobs, each containing a template path, an output path, and a DataTable. For each job it loads the template (or creates a blank workbook if the file is missing), initializes a WorkbookDesigner, sets the job‑specific data source, processes smart markers, and saves the result. Sample product and employee tables illustrate distinct data sources per workbook.
class BatchSmartMarkerProcessor
{
    static void Main()
    {
        // Define jobs: template path, output path and associated data source
        var jobs = new List<(string templatePath, string outputPath, DataTable dataSource)>
        {
            ("Template1.xlsx", "Result1.xlsx", CreateProductsTable()),
            ("Template2.xlsx", "Result2.xlsx", CreateEmployeesTable())
        };

        foreach (var job in jobs)
        {
            try
            {
                Workbook templateWorkbook;

                // Load existing template or create a new blank workbook if the file is missing
                if (File.Exists(job.templatePath))
                {
                    templateWorkbook = new Workbook(job.templatePath);
                }
                else
                {
                    Console.WriteLine($"Template file not found: {job.templatePath}. Creating a blank workbook.");
                    templateWorkbook = new Workbook(); // creates a default workbook with one worksheet
                }

                // Initialize designer with the workbook
                WorkbookDesigner designer = new WorkbookDesigner(templateWorkbook);

                // Assign the data source for this job
                designer.SetDataSource(job.dataSource);

                // Process smart markers
                designer.Process();

                // Save the result
                designer.Workbook.Save(job.outputPath);
                Console.WriteLine($"Processed and saved: {job.outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing job for template '{job.templatePath}': {ex.Message}");
            }
        }
    }

    // Sample data source for the first template
    static DataTable CreateProductsTable()
    {
        DataTable dt = new DataTable("Products");
        dt.Columns.Add("ProductID", typeof(int));
        dt.Columns.Add("ProductName", typeof(string));
        dt.Columns.Add("Price", typeof(decimal));

        dt.Rows.Add(1, "Apple", 0.5m);
        dt.Rows.Add(2, "Banana", 0.3m);
        dt.Rows.Add(3, "Cherry", 0.8m);

        return dt;
    }

    // Sample data source for the second template
    static DataTable CreateEmployeesTable()
    {
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("EmployeeID", typeof(int));
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Department", typeof(string));

        dt.Rows.Add(101, "John Doe", "Sales");
        dt.Rows.Add(102, "Jane Smith", "HR");
        dt.Rows.Add(103, "Mike Johnson", "IT");

        return dt;
    }
}
