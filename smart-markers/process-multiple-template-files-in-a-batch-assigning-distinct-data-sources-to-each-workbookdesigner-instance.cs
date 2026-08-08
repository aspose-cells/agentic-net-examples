// Title: Batch process multiple Excel templates with distinct data sources using Aspose.Cells WorkbookDesigner (C#)
// Description: Iterates through a list of Excel template files, creates a dedicated DataTable for each, loads or creates the workbook, binds the appropriate table to WorkbookDesigner, processes smart markers, and saves the result. Includes graceful handling of missing templates and runtime errors.
// Keywords: Aspose.Cells | WorkbookDesigner | C# smart markers | batch Excel generation | multiple templates | SetDataSource | Excel report automation | DataTable binding | template processing loop | error handling
// Common Searches: Aspose.Cells batch processing multiple templates C# | WorkbookDesigner SetDataSource for each Excel file | how to generate several Excel reports from different data tables | smart markers loop over templates Aspose.Cells | C# code to process multiple Excel templates with WorkbookDesigner
// Developer Intent: Automatically generate a series of Excel reports by applying a unique DataTable to each corresponding template workbook using WorkbookDesigner.
// Use Cases: Create an employee directory by binding the Employees DataTable to Template1.xlsx and saving Result1.xlsx. | Produce a product catalog by applying the Products DataTable to Template2.xlsx and outputting Result2.xlsx. | Generate an order summary by linking the Orders DataTable to Template3.xlsx and saving Result3.xlsx.
// AI Prompts: Add support for a fourth template with its own DataTable and output file. | Implement detailed logging for each workbook's processing steps and errors. | Customize the smart marker delimiters (e.g., {{}} instead of &) for all templates in the batch.

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace BatchWorkbookDesignerDemo
{
    // Iterates through a list of Excel template files, creates a dedicated DataTable for each, loads or creates the workbook, binds the appropriate table to WorkbookDesigner, processes smart markers, and saves the result. Includes graceful handling of missing templates and runtime errors.
    class Program
    {
        static void Main()
        {
            // Define template files and corresponding output files
            var templates = new[]
            {
                new { TemplatePath = "Template1.xlsx", OutputPath = "Result1.xlsx" },
                new { TemplatePath = "Template2.xlsx", OutputPath = "Result2.xlsx" },
                new { TemplatePath = "Template3.xlsx", OutputPath = "Result3.xlsx" }
            };

            // Prepare distinct data sources for each template
            var dataSources = new List<DataTable>();

            // Data source for Template1
            var dt1 = new DataTable("Employees");
            dt1.Columns.Add("Name", typeof(string));
            dt1.Columns.Add("Age", typeof(int));
            dt1.Rows.Add("John Doe", 30);
            dt1.Rows.Add("Jane Smith", 28);
            dataSources.Add(dt1);

            // Data source for Template2
            var dt2 = new DataTable("Products");
            dt2.Columns.Add("ProductID", typeof(int));
            dt2.Columns.Add("ProductName", typeof(string));
            dt2.Columns.Add("Price", typeof(decimal));
            dt2.Rows.Add(101, "Laptop", 1200.50m);
            dt2.Rows.Add(102, "Smartphone", 799.99m);
            dataSources.Add(dt2);

            // Data source for Template3
            var dt3 = new DataTable("Orders");
            dt3.Columns.Add("OrderID", typeof(int));
            dt3.Columns.Add("Customer", typeof(string));
            dt3.Columns.Add("Total", typeof(decimal));
            dt3.Rows.Add(5001, "Acme Corp", 2500.00m);
            dt3.Rows.Add(5002, "Globex Inc", 1800.75m);
            dataSources.Add(dt3);

            // Process each template with its specific data source
            for (int i = 0; i < templates.Length; i++)
            {
                try
                {
                    Workbook workbook;

                    // Load the template workbook if it exists; otherwise create a blank workbook
                    if (File.Exists(templates[i].TemplatePath))
                    {
                        workbook = new Workbook(templates[i].TemplatePath);
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Template file '{templates[i].TemplatePath}' not found. Creating a blank workbook.");
                        workbook = new Workbook(); // creates a new empty workbook
                    }

                    // Initialize WorkbookDesigner with the loaded workbook
                    var designer = new WorkbookDesigner(workbook);

                    // Assign the distinct data source (using the table name as the data source name)
                    // The smart markers in the template should reference this name, e.g., &Employees.Name
                    designer.SetDataSource(dataSources[i].TableName, dataSources[i]);

                    // Process the smart markers
                    designer.Process();

                    // Save the processed workbook to the specified output file
                    designer.Workbook.Save(templates[i].OutputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing template '{templates[i].TemplatePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
