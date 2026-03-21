using System;
using Aspose.Cells;

namespace SmartMarkersDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the template workbook that contains smart markers (e.g., &=$Employee.Name)
            string templatePath = "TemplateWithSmartMarkers.xlsx";

            // Path where the generated workbook will be saved
            string outputPath = "GeneratedReport.xlsx";

            // Load the template workbook (create rule)
            Workbook workbook = new Workbook(templatePath);

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // JSON data source that matches the smart marker groups in the template
            string jsonData = @"{
                ""Employee"": {
                    ""Name"": ""Alice Johnson"",
                    ""Age"": 29,
                    ""Department"": ""Finance"",
                    ""Salary"": 72000.50
                },
                ""Projects"": [
                    { ""Title"": ""Budget Analysis"", ""Hours"": 120 },
                    { ""Title"": ""Audit Preparation"", ""Hours"": 80 }
                ]
            }";

            // Bind the JSON data to the smart marker groups
            // Smart markers like &=$Employee.Name will be replaced with the corresponding values
            designer.SetJsonDataSource("Employee", jsonData);
            // Smart markers like &=$Projects.Title will iterate over the Projects array
            designer.SetJsonDataSource("Projects", jsonData);

            // Process all smart markers in the workbook (process rule)
            designer.Process();

            // Save the populated workbook (save rule)
            workbook.Save(outputPath);
        }
    }
}