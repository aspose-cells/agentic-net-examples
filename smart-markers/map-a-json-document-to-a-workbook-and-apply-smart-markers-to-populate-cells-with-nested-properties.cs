// Title: Map nested JSON data to an Excel workbook using Aspose.Cells smart markers with WorkbookDesigner in C#
// Common Searches: C# example of using Aspose.Cells WorkbookDesigner to fill an Excel template with JSON data | How to flatten nested JSON objects for smart markers in Aspose.Cells | Populate an Excel file from a DataSet using smart markers in C# | Aspose.Cells smart markers handling nested address fields in a template
// Tags: Aspose.Cells WorkbookDesigner set DataSet source | smart markers populate Excel from DataTable C# | flatten nested JSON for Excel export Aspose.Cells | Excel template smart markers with nested objects | generate Excel workbook from JSON using Aspose.Cells C#

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // for WorkbookDesigner

// The example loads an Excel template containing smart markers, manually builds a DataSet that flattens nested address fields from a JSON structure, assigns the DataSet to a WorkbookDesigner, processes the smart markers to populate the worksheet, ensures the output directory exists, and saves the resulting workbook.
class JsonToWorkbook
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the Excel template that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // JSON data (kept for reference; not used directly in this example)
            string jsonData = @"
            {
                ""Employees"": [
                    {
                        ""Name"": ""John"",
                        ""Address"": {
                            ""City"": ""New York"",
                            ""Zip"": ""10001""
                        }
                    },
                    {
                        ""Name"": ""Anna"",
                        ""Address"": {
                            ""City"": ""London"",
                            ""Zip"": ""SW1A""
                        }
                    }
                ]
            }";

            // Build a DataSet manually (flattening nested address fields)
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("City", typeof(string));
            employeeTable.Columns.Add("Zip", typeof(string));

            employeeTable.Rows.Add("John", "New York", "10001");
            employeeTable.Rows.Add("Anna", "London", "SW1A");

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(employeeTable);

            // Initialize the designer and set the data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);

            // Process the smart markers and populate the worksheet
            designer.Process();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the populated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
