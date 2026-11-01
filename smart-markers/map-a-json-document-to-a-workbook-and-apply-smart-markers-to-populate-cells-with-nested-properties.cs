// Title: Use Aspose.Cells WorkbookDesigner to populate an Excel template with nested JSON data (employees with address) in C#
// AI Prompts: Write C# code that loads an .xlsx template, sets a JSON string containing an Employees array with nested Address objects as the data source for WorkbookDesigner, processes smart markers, and saves the resulting workbook. | Show the exact smart‑marker syntax to reference nested JSON fields such as Address.Street and Address.City inside the Excel template. | Adapt the sample to read the JSON payload from an external file, handle missing template errors, and still generate the populated workbook.
// Common Searches: asp.net c# bind hierarchical json to excel using aspose.cells smart markers | populate excel template from json array with nested objects using workbookdesigner | aspose.cells smart markers reference nested property address city in c# | load excel template set json data source process smart markers example | c# read json file and fill excel template with employee address using aspose.cells
// Tags: Aspose.Cells WorkbookDesigner JSON data source | smart markers nested object mapping | populate Excel template from hierarchical JSON C# | reference Address.Street in smart marker | load template and save result with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads an Excel template, assigns a JSON string containing an Employees array with nested Address objects as the data source for WorkbookDesigner, processes smart markers to fill the worksheet, and saves the populated workbook as Result.xlsx.
class JsonToWorkbook
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";

            // Ensure the template file exists before loading
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the Excel template that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // JSON data with nested properties
            string jsonData = @"
            {
                ""Employees"": [
                    {
                        ""Name"": ""John"",
                        ""Address"": {
                            ""Street"": ""123 Main St"",
                            ""City"": ""New York""
                        }
                    },
                    {
                        ""Name"": ""Jane"",
                        ""Address"": {
                            ""Street"": ""456 Oak Ave"",
                            ""City"": ""Los Angeles""
                        }
                    }
                ]
            }";

            // Create a WorkbookDesigner to work with smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the JSON string as the data source for the designer
            // Use the overload that accepts JSON data and a data source name
            designer.SetDataSource(jsonData, "json");

            // Process the smart markers and populate the worksheet
            designer.Process();

            // Save the populated workbook
            const string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
