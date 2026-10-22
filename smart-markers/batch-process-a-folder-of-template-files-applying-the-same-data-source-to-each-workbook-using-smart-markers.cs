// Title: How to batch process Excel template files with smart markers using a shared DataTable in Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates over all .xlsx files in a given folder, assigns a common DataTable to WorkbookDesigner, processes the smart markers, and saves each workbook to a target directory. | Show how to create a sample DataTable and use it as the data source for smart markers across multiple workbooks in a single batch operation with Aspose.Cells. | Provide a robust method that validates input and output folders, loads each template workbook, applies the shared data source, processes smart markers, and includes error handling for batch processing.
// Common Searches: Aspose.Cells batch processing smart markers from a single DataTable in C# | C# loop through folder of .xlsx templates and populate smart markers with the same data source | How to use WorkbookDesigner to apply one DataTable to many Excel files using Aspose.Cells | Automate bulk smart marker replacement in multiple Excel workbooks with Aspose.Cells .NET | Validate source and destination directories when processing smart markers in batch with Aspose.Cells
// Tags: batch smart marker processing Aspose.Cells | WorkbookDesigner set DataTable for multiple workbooks | process multiple .xlsx templates with smart markers | shared data source for Excel smart markers .NET | automated Excel template population Aspose.Cells

using System;
using System.IO;
using System.Data;
using Aspose.Cells;

namespace BatchSmartMarkerProcessing
{
    // The example scans a folder for .xlsx template files, loads each workbook, assigns a shared DataTable as the smart‑marker data source via WorkbookDesigner, processes the markers, and saves the populated workbooks to an output directory, handling folder validation and errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Folder containing template workbooks with smart markers
                string templateFolder = @"C:\Templates";
                // Folder where processed workbooks will be saved
                string outputFolder = @"C:\Processed";

                // Verify template folder exists
                if (!Directory.Exists(templateFolder))
                {
                    Console.WriteLine($"Template folder does not exist: {templateFolder}");
                    return;
                }

                // Ensure output folder exists
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                // Create a sample data source (DataTable) to be applied to all workbooks
                DataTable dataSource = GetSampleData();

                // Process each .xlsx file in the template folder
                foreach (string templatePath in Directory.GetFiles(templateFolder, "*.xlsx"))
                {
                    // Verify the template file exists
                    if (!File.Exists(templatePath))
                    {
                        Console.WriteLine($"File not found: {templatePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(templatePath);

                    // Initialize the WorkbookDesigner for smart marker processing
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);

                    // Set the data source for smart markers
                    designer.SetDataSource(dataSource);

                    // Process the smart markers in the workbook
                    designer.Process();

                    // Build the output file path
                    string fileName = Path.GetFileNameWithoutExtension(templatePath);
                    string outputPath = Path.Combine(outputFolder, $"{fileName}_Processed.xlsx");

                    // Save the processed workbook
                    workbook.Save(outputPath);
                }

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Generates a sample DataTable used as the data source for smart markers
        private static DataTable GetSampleData()
        {
            DataTable table = new DataTable("Employees");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Department", typeof(string));
            table.Columns.Add("Salary", typeof(double));

            table.Rows.Add("John Doe", "Finance", 75000);
            table.Rows.Add("Jane Smith", "HR", 68000);
            table.Rows.Add("Mike Johnson", "IT", 82000);
            table.Rows.Add("Emily Davis", "Marketing", 71000);

            return table;
        }
    }
}
