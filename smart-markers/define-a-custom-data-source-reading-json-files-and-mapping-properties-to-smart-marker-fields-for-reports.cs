// Title: Generate an Excel report from a JSON file using Aspose.Cells smart markers in C#
// AI Prompts: Write C# code that loads an Excel template, reads a JSON file, assigns the JSON to a smart‑marker named 'Employee' with WorkbookDesigner.SetJsonDataSource, processes the markers, and saves the output workbook. | Modify the sample to handle a JSON array of employees and bind it to a table smart marker, including validation for missing template or data files. | Create a reusable C# method that accepts template path, JSON string, and output path, then uses WorkbookDesigner to populate smart markers and return the generated workbook.
// Common Searches: how to bind JSON data to Aspose.Cells smart markers in C# | Aspose.Cells WorkbookDesigner SetJsonDataSource example for Excel report | populate Excel template with JSON array using Aspose.Cells .NET | C# generate Excel file from JSON using smart markers Aspose | error handling for missing template or JSON file Aspose.Cells
// Tags: Aspose.Cells SetJsonDataSource smart markers | C# fill Excel template from JSON | WorkbookDesigner smart marker processing | JSON-driven Excel report generation | Excel smart marker data binding .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonSmartMarkerDemo
{
    // The example verifies the presence of an Excel template and a JSON file, loads the template into a Workbook, reads the JSON content, assigns it to the 'Employee' smart marker via WorkbookDesigner.SetJsonDataSource, processes all smart markers to populate the worksheet, and saves the resulting workbook as Report_Output.xlsx while handling potential exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the Excel template that contains smart markers, e.g. &=$Employee.Name
                string templatePath = "Template.xlsx";

                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook
                Workbook workbook = new Workbook(templatePath);

                // Create a WorkbookDesigner and assign the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Path to the JSON data file
                string jsonFilePath = "EmployeeData.json";

                // Verify that the JSON file exists
                if (!File.Exists(jsonFilePath))
                {
                    Console.WriteLine($"JSON data file not found: {jsonFilePath}");
                    return;
                }

                // Read JSON content from the file
                string jsonData = File.ReadAllText(jsonFilePath);

                // Set the JSON string as a data source for the smart marker named "Employee"
                designer.SetJsonDataSource("Employee", jsonData);

                // Process the smart markers and populate the worksheet with data
                designer.Process();

                // Save the resulting workbook
                string outputPath = "Report_Output.xlsx";
                designer.Workbook.Save(outputPath);

                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
