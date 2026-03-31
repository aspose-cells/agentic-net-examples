using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonSmartMarkerDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel template that contains smart markers (e.g., &=$DataSource.Name)
            string templatePath = "Template.xlsx";

            // Path where the populated workbook will be saved
            string outputPath = "Result.xlsx";

            // Sample JSON data that matches the smart markers in the template
            // Example smart marker in the template: &=$DataSource.Name
            string jsonData = @"{
                ""Name"": ""John Doe"",
                ""Age"": 30,
                ""City"": ""New York""
            }";

            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the JSON string as a data source for smart markers.
            // The first parameter is the name used in the smart marker (DataSource in this case).
            designer.SetJsonDataSource("DataSource", jsonData);

            // Process all smart markers in the workbook and populate them with JSON data
            designer.Process();

            // Save the populated workbook to the specified output file
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook has been generated and saved to '{outputPath}'.");
        }
    }
}