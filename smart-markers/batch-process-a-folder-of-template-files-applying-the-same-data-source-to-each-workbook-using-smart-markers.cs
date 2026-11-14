// Title: Batch fill multiple Excel template workbooks with a shared DataTable using Aspose.Cells smart markers in C#
// Common Searches: c# how to use WorkbookDesigner to apply one DataTable to all Excel templates in a folder | batch fill .xlsx templates with smart markers using Aspose.Cells | process multiple Excel workbooks with the same smart‑marker data source in C# | apply shared smart marker data to a directory of Excel files with Aspose.Cells | automate filling of many Excel templates from a single DataTable using smart markers
// Tags: WorkbookDesigner batch processing smart markers | shared DataTable smart marker source Aspose.Cells | populate multiple .xlsx templates C# | folder-wide Excel template filling using smart markers | bulk smart marker processing Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The program scans a specified input folder for .xlsx template files, loads each workbook, assigns a single DataTable as the smart‑marker data source via WorkbookDesigner, processes the smart markers, and saves the populated workbooks to an output folder with modified filenames.
class Program
{
    static void Main()
    {
        try
        {
            // Folder containing the template workbooks
            string inputFolder = @"C:\Templates";

            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\Processed";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Create a single data source that will be applied to all workbooks
            DataTable dataSource = CreateDataSource();

            // Iterate over each Excel template in the input folder
            foreach (string templatePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Verify the template file exists before attempting to load
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (load rule)
                    Workbook workbook = new Workbook(templatePath);

                    // Process smart markers with the shared data source (smart‑marker rule)
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);
                    designer.SetDataSource(dataSource);
                    designer.Process();

                    // Build the output file name
                    string fileName = Path.GetFileNameWithoutExtension(templatePath);
                    string outputPath = Path.Combine(outputFolder, $"{fileName}_filled.xlsx");

                    // Save the processed workbook (save rule)
                    workbook.Save(outputPath);
                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{templatePath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Generates a sample DataTable used as the smart‑marker data source
    static DataTable CreateDataSource()
    {
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Columns.Add("Department", typeof(string));

        table.Rows.Add("John Doe", 30, "Sales");
        table.Rows.Add("Jane Smith", 28, "Marketing");
        table.Rows.Add("Bob Johnson", 35, "IT");

        return table;
    }
}
