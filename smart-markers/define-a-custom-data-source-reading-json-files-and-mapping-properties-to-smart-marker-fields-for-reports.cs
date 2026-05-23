using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonSmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the Excel template that contains smart markers, e.g. "&=$Employee.Name"
                string templatePath = "TemplateWithSmartMarkers.xlsx";

                // Path to the JSON file that holds the data source
                string jsonFilePath = "Employees.json";

                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Verify that the JSON file exists
                if (!File.Exists(jsonFilePath))
                {
                    Console.WriteLine($"JSON data file not found: {jsonFilePath}");
                    return;
                }

                // Load the JSON content from file
                string jsonData = File.ReadAllText(jsonFilePath);

                // Load the workbook template
                Workbook workbook = new Workbook(templatePath);

                // Create a WorkbookDesigner and assign the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Set the JSON string as a data source named "Employee"
                // Smart markers in the template should reference this name, e.g. &=$Employee.Name
                designer.SetJsonDataSource("Employee", jsonData);

                // Process all smart markers in the workbook
                designer.Process();

                // Save the populated workbook
                string outputPath = "Report_Output.xlsx";
                designer.Workbook.Save(outputPath);

                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File not found: {ex.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}