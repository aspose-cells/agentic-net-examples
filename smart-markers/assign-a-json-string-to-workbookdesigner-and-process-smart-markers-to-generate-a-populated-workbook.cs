// Title: C# – Populate Excel with Aspose.Cells WorkbookDesigner using JSON smart markers
// Description: Demonstrates how to create a workbook, add smart markers, bind a JSON string to the WorkbookDesigner via SetJsonDataSource, process the markers, and save the populated Excel file (EmployeeReport.xlsx) using Aspose.Cells.
// Keywords: Aspose.Cells | WorkbookDesigner | smart markers | JSON data source | C# Excel generation | SetJsonDataSource example | populate Excel from JSON | Aspose.Cells tutorial
// Common Searches: Aspose.Cells WorkbookDesigner JSON example | C# smart markers with JSON data | How to bind JSON to Excel using Aspose.Cells | SetJsonDataSource usage in .NET | Generate Excel report from JSON Aspose
// Developer Intent: Assign a JSON string to WorkbookDesigner and process smart markers to generate a populated workbook.
// Use Cases: Create a quick employee report by inserting smart markers and filling them with JSON data. | Transform API responses into formatted Excel files without manual cell mapping. | Automate dynamic Excel generation in a .NET backend using smart markers and JSON sources.
// AI Prompts: Write C# code that binds a JSON array of employee objects to smart markers and expands rows automatically. | Explain the relationship between the marker prefix and SetJsonDataSource, and show how to handle nested JSON objects. | Provide a step‑by‑step guide to load a JSON file, assign it to WorkbookDesigner, process smart markers, and save the workbook with styling.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add smart markers, bind a JSON string to the WorkbookDesigner via SetJsonDataSource, process the markers, and save the populated Excel file (EmployeeReport.xlsx) using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert smart markers that reference the JSON data source
        sheet.Cells["A1"].PutValue("&=$Employee.Name");
        sheet.Cells["B1"].PutValue("&=$Employee.Age");
        sheet.Cells["C1"].PutValue("&=$Employee.City");

        // Initialize the WorkbookDesigner with the workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // JSON string that will be used as the data source
        string json = "{\"Name\":\"John Doe\",\"Age\":30,\"City\":\"New York\"}";

        // Assign the JSON string to the designer (name must match the marker prefix)
        designer.SetJsonDataSource("Employee", json);

        // Process the smart markers to populate the worksheet
        designer.Process();

        // Save the resulting workbook
        workbook.Save("EmployeeReport.xlsx");
    }
}
