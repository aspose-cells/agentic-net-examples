using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace BatchWorkbookDesignerDemo
{
    public class BatchProcessor
    {
        /// <summary>
        /// Processes a batch of Excel template files. Each template is bound to its own data source,
        /// processed with smart markers, and saved to a corresponding output file.
        /// </summary>
        /// <param name="templateFiles">Full paths of the template Excel files.</param>
        /// <param name="dataSources">Data sources (DataTable, List<T>, etc.) for each template.</param>
        /// <param name="outputFiles">Full paths where the processed workbooks will be saved.</param>
        public void ProcessBatch(IList<string> templateFiles, IList<object> dataSources, IList<string> outputFiles)
        {
            if (templateFiles == null) throw new ArgumentNullException(nameof(templateFiles));
            if (dataSources == null) throw new ArgumentNullException(nameof(dataSources));
            if (outputFiles == null) throw new ArgumentNullException(nameof(outputFiles));
            if (templateFiles.Count != dataSources.Count || templateFiles.Count != outputFiles.Count)
                throw new ArgumentException("All input collections must have the same number of elements.");

            for (int i = 0; i < templateFiles.Count; i++)
            {
                try
                {
                    // Verify that the template file exists
                    if (!File.Exists(templateFiles[i]))
                        throw new FileNotFoundException($"Template file not found: {templateFiles[i]}");

                    // Load the template workbook
                    Workbook workbook = new Workbook(templateFiles[i]);

                    // Assign workbook to a designer
                    WorkbookDesigner designer = new WorkbookDesigner { Workbook = workbook };

                    // Bind the data source to the variable name "Data"
                    designer.SetDataSource("Data", dataSources[i]);

                    // Process smart markers
                    designer.Process();

                    // Ensure output directory exists
                    string outputDir = Path.GetDirectoryName(outputFiles[i]);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                        Directory.CreateDirectory(outputDir);

                    // Save the processed workbook
                    designer.Workbook.Save(outputFiles[i]);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file pair #{i + 1}: {ex.Message}");
                    // Optionally continue with next file or rethrow
                }
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare a list of template files
                var templates = new List<string>
                {
                    @"C:\Templates\Report1.xlsx",
                    @"C:\Templates\Report2.xlsx"
                };

                // Prepare corresponding data sources
                var dataSources = new List<object>();

                // First data source: a DataTable
                DataTable dt1 = new DataTable("Sales");
                dt1.Columns.Add("Region", typeof(string));
                dt1.Columns.Add("Amount", typeof(decimal));
                dt1.Rows.Add("North", 1200.50m);
                dt1.Rows.Add("South", 950.75m);
                dataSources.Add(dt1);

                // Second data source: a list of custom objects
                var employees = new List<Employee>
                {
                    new Employee { Name = "John Doe", Age = 30 },
                    new Employee { Name = "Jane Smith", Age = 28 }
                };
                dataSources.Add(employees);

                // Prepare output file paths
                var outputs = new List<string>
                {
                    @"C:\Outputs\Report1_Processed.xlsx",
                    @"C:\Outputs\Report2_Processed.xlsx"
                };

                // Execute batch processing
                var processor = new BatchProcessor();
                processor.ProcessBatch(templates, dataSources, outputs);

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }

    // Sample custom class used in the second data source
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}