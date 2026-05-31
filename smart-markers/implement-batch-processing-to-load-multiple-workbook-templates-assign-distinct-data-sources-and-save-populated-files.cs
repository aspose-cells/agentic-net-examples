using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    // Represents a single batch item: template file, data source name, data object, and output file.
    public class BatchItem
    {
        public string TemplatePath { get; set; }      // Path to the template workbook (contains smart markers)
        public string DataSourceName { get; set; }    // Name used in smart markers, e.g. "&DataSourceName.Property"
        public object DataSource { get; set; }        // The actual data object (List<T>, DataTable, etc.)
        public string OutputPath { get; set; }        // Where the populated workbook will be saved
    }

    public static class BatchProcessor
    {
        // Processes a collection of BatchItem objects.
        public static void ProcessTemplates(IEnumerable<BatchItem> items)
        {
            foreach (var item in items)
            {
                try
                {
                    // Verify that the template file exists.
                    if (!File.Exists(item.TemplatePath))
                        throw new FileNotFoundException($"Template file not found: {item.TemplatePath}");

                    // Load the template workbook.
                    var templateWorkbook = new Workbook(item.TemplatePath);

                    // Initialize WorkbookDesigner with the loaded workbook.
                    var designer = new WorkbookDesigner(templateWorkbook);

                    // Assign the distinct data source to the designer.
                    designer.SetDataSource(item.DataSourceName, item.DataSource);

                    // Process smart markers and populate the workbook.
                    designer.Process();

                    // Ensure the output directory exists.
                    var outputDir = Path.GetDirectoryName(item.OutputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                        Directory.CreateDirectory(outputDir);

                    // Save the populated workbook to the specified output path.
                    designer.Workbook.Save(item.OutputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing batch item (Template: {item.TemplatePath}, Output: {item.OutputPath}): {ex.Message}");
                }
            }
        }
    }

    // Example usage.
    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare distinct data sources for each template.
                var employees = new List<Employee>
                {
                    new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
                    new Employee { Name = "Jane Smith", Age = 28, Department = "HR" }
                };

                var products = new List<Product>
                {
                    new Product { Id = 101, Name = "Laptop", Price = 1200.00 },
                    new Product { Id = 102, Name = "Smartphone", Price = 799.99 }
                };

                // Define batch items.
                var batchItems = new List<BatchItem>
                {
                    new BatchItem
                    {
                        TemplatePath = "EmployeeTemplate.xlsx",
                        DataSourceName = "Employees",
                        DataSource = employees,
                        OutputPath = "EmployeeReport.xlsx"
                    },
                    new BatchItem
                    {
                        TemplatePath = "ProductTemplate.xlsx",
                        DataSourceName = "Products",
                        DataSource = products,
                        OutputPath = "ProductCatalog.xlsx"
                    }
                };

                // Execute batch processing.
                BatchProcessor.ProcessTemplates(batchItems);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }

    // Sample data classes used in the example.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }

    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public double Price { get; set; }
    }
}