// Title: Create an Excel report with Aspose.Cells .NET using JSON smart markers
// Description: Loads an Excel template, inserts smart‑marker placeholders, reads a JSON file with employee data, assigns the JSON as a data source to WorkbookDesigner, processes the markers, and saves the populated workbook.
// Keywords: Aspose.Cells | JSON | smart markers | WorkbookDesigner | C# | .NET | Excel template | data binding | populate worksheet | report generation
// Common Searches: Aspose.Cells bind JSON to smart markers | C# set JSON data source for WorkbookDesigner | Generate Excel from JSON using smart markers | How to use smart markers with JSON in Aspose.Cells | Load JSON file into Excel template Aspose.Cells
// Developer Intent: Bind JSON data to smart‑marker fields in an Excel template with Aspose.Cells for .NET.
// Use Cases: Populate employee name and age columns by mapping a JSON employee list to smart markers. | Create a sales dashboard where JSON sales records fill detail rows and summary cells via smart markers. | Automate invoice generation by reading JSON order information and inserting customer and line‑item data into a pre‑designed template.
// AI Prompts: Show me C# code that reads a JSON file and sets it as a data source for WorkbookDesigner in Aspose.Cells. | How can I add smart‑marker placeholders programmatically and process them with JSON data in .NET? | Explain handling of nested JSON objects with smart markers using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel template, inserts smart‑marker placeholders, reads a JSON file with employee data, assigns the JSON as a data source to WorkbookDesigner, processes the markers, and saves the populated workbook.
    class JsonSmartMarkerExample
    {
        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string jsonPath = "employees.json";
                const string outputPath = "Report.xlsx";

                // Verify that the template file exists.
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook.
                Workbook workbook = new Workbook(templatePath);

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Add smart markers (if they are not already present).
                sheet.Cells["A1"].PutValue("&=$Employee.Name");
                sheet.Cells["B1"].PutValue("&=$Employee.Age");

                // Create and configure the designer.
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Verify that the JSON data file exists.
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"JSON data file not found: {jsonPath}");
                    return;
                }

                // Read JSON data.
                string json = File.ReadAllText(jsonPath);

                // Set the JSON data source for the smart markers.
                designer.SetJsonDataSource("Employee", json);

                // Process all smart markers.
                designer.Process();

                // Save the populated workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
