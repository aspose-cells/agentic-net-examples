// Title: Batch fill multiple Excel templates with a shared DataTable using Aspose.Cells smart markers (C#)
// Description: C# sample that scans a folder for .xlsx templates, loads each workbook with WorkbookDesigner, assigns a common DataTable to the smart‑marker prefix, processes all smart markers, and saves the populated files to an output directory while handling errors.
// Keywords: Aspose.Cells batch processing | C# smart markers | WorkbookDesigner data source | populate multiple Excel templates | shared DataTable Excel | process folder of .xlsx files
// Common Searches: Aspose.Cells batch smart marker example C# | apply one DataTable to many Excel templates | process all .xlsx files in a folder with smart markers | WorkbookDesigner set data source for multiple workbooks | C# fill Excel templates from a folder
// Developer Intent: Use a single DataTable to populate smart markers in every Excel template within a directory and save the results.
// Use Cases: Generate department‑specific reports by applying the same employee DataTable to several template workbooks. | Create a batch of invoices where a common customer list is merged into each invoice template via smart markers. | Automate monthly dashboards by populating multiple template files with identical metric data stored in a DataTable.
// AI Prompts: Write C# code that loads all .xlsx files from a directory, sets a shared DataTable as the data source for smart markers using WorkbookDesigner, processes them, and saves the output to another folder. | Show how to add robust error handling while batch processing Excel templates with Aspose.Cells smart markers. | Explain how to extend the example to use multiple DataTables with different smart‑marker prefixes in a single batch operation.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace BatchSmartMarkerProcessing
{
    // C# sample that scans a folder for .xlsx templates, loads each workbook with WorkbookDesigner, assigns a common DataTable to the smart‑marker prefix, processes all smart markers, and saves the populated files to an output directory while handling errors.
    class Program
    {
        static void Main()
        {
            // Folder containing template workbooks with smart markers
            string templatesFolder = @"C:\Templates";
            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\Processed";

            // Ensure the output folder exists
            Directory.CreateDirectory(outputFolder);

            // Verify that the templates folder exists
            if (!Directory.Exists(templatesFolder))
            {
                Console.WriteLine($"Templates folder not found: {templatesFolder}");
                return;
            }

            // Prepare a common data source (DataTable) that will be applied to every workbook
            DataTable commonData = CreateSampleDataTable();

            // Get all Excel files in the templates folder
            string[] templateFiles = Directory.GetFiles(templatesFolder, "*.xlsx");

            if (templateFiles.Length == 0)
            {
                Console.WriteLine("No template files found.");
                return;
            }

            foreach (string templatePath in templateFiles)
            {
                try
                {
                    // Ensure the template file exists before loading
                    if (!File.Exists(templatePath))
                    {
                        Console.WriteLine($"File not found: {templatePath}");
                        continue;
                    }

                    // Load the template workbook
                    Workbook workbook = new Workbook(templatePath);

                    // Initialize WorkbookDesigner with the loaded workbook
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);

                    // Set the common data source; the name "Data" must match the smart marker prefix in the templates
                    designer.SetDataSource("Data", commonData);

                    // Process all smart markers in the workbook
                    designer.Process();

                    // Build the output file path (same file name, different folder)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(templatePath));

                    // Save the processed workbook
                    workbook.Save(outputPath);

                    Console.WriteLine($"Processed: {Path.GetFileName(templatePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{templatePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }

        // Helper method to create a sample DataTable used as the common data source
        private static DataTable CreateSampleDataTable()
        {
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Department", typeof(string));

            dt.Rows.Add("John Doe", 30, "Sales");
            dt.Rows.Add("Jane Smith", 28, "Marketing");
            dt.Rows.Add("Bob Johnson", 35, "Engineering");

            return dt;
        }
    }
}
